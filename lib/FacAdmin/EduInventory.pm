# -*- cperl -*-
package FacAdmin::EduInventory;

=head1 NAME

FacAdmin::EduInventory -- Inventory of all faculty teaching data based on XLDB

=cut


require 5.006;
use strict;
use parent 'XLDB::DataBase';
use ConLogger;
use ConLogger::ProgressBar;

# constructor 'new' inherited from parent

# build the database
sub buildDataBase {
  my $self = shift;

  my ( $pers ) = @_;

  ConLogger::logitem( "Reading header" );
  $self->_readHeader( 'Studiegidsnr',
		      {
		       '^Studiegidsnr$'              => 'SGNR',
		       '^Curriculum$'                => 'CURRICULUM',
		       '^Rooster$'                   => 'ROOSTER',
		       '^Opdracht$'                  => 'OPDRACHT',
		       '^Programmajaar$'             => 'PROGJAAR',
		       '^Programma '                 => 'PROG',
		       '^Deeltijds$'                 => 'DEELTIJDS',
		       '^Ba-(.*)$'                   => 'BA-',
		       '^Ma-(.*)$'                   => 'MA-',
		       '^Minor - (.*)$'              => 'MINOR-',
		       '^Afstudeerrichting (.)$'     => 'AR-',
		       '^Keuzepakket'                => 'KP',
		       '^SP-(.*)$'                   => 'SP-',
		       '^VT-(.*)$'                   => 'VT-',
		       '^VP-(.*)$'                   => 'VP-',
		       '^Categorie$'                 => 'CAT',
		       '^Opleidingsonderdeel$'       => 'OO',
		       '^Te roosteren activiteit$'   => 'ACT',
		       '^Groep'                      => 'GRP',
		       '^SP$$'                       => 'SP',
		       '^Hoorcollege'                => 'HC',
		       '^Practicum'                  => 'PR',
		       '^Werkcollege / Seminarie'    => 'WC',
		       '^Coaching / PGO'             => 'CP',
		       '^Wetenschappelijk project'   => 'WP',
		       '^Bachelorproef'              => 'BP',
		       '^Stage'                      => 'ST',
		       '^Masterproef'                => 'MP',
		       '^Toeslag grote groepen'      => 'GGR',
		       '^Aantal groepen'             => 'AGR',
		       '^Nieuwe activiteit'          => 'NEW',
		       '^Titularis'                  => 'TITULARIS',
		       '^Docent'                     => 'DOCENT',
		       '^Semester'                   => 'SEM',
		       '^Aantal sessies$'            => 'ASS',
		       '^Duur 1 sessie'              => 'DSS',
		       '^Totaal aantal roosteruren'  => 'TOTAAL',
		       '^Campus'                     => 'CAMPUS',
		       '^Lokaal$'                    => 'LOKAAL',
		       '^Opmerking'                  => 'OPM',
		      }
		    );

  ConLogger::logitem( "Reading data records" );

  {
    my $pbar = ConLogger::ProgressBar->new();

    for ( my $i = $self->property('HeaderRow')+1; $i <= $self->property('MaxRow'); ++$i ) {

      $pbar->progress( $i, $self->property('MaxRow') );

      ######################################
      # parse the line into the %$line hash
      my $line = $self->parseLine( $i );

      ###########################
      # do generic sanity checks
      # check presence of mandatory fields
      $self-> checkMandatoryFields( [ qw( SGNR PROGJAAR OO SEM ) ], $line, "" );

      # check field values
      $self->checkEnumFields( [ qw( ^SEM$ ) ], [ qw( 1 2 12 ) ], $line,  );
      $self->checkPositiveIntegerFields( [ qw( PROGJAAR ST MP ) ], $line );
      $self->checkPositiveRealFields( [ qw( HC PR WC CP WP BP ) ], $line );

      # check validity of studiegidsnr
      dieOnCurrentLine( "invalid '$self->{header}->{SGNR}'-value '$line->{SGNR}'\n" )
	unless ( $line->{SGNR} =~ /^\d\d\d\dFTI\w\w\w$/ );

      # check true or false columns
      # this also completes empty true or false cells
      $self->checkBooleanFields( [ qw( ^CURRICULUM$ ^ROOSTER$ ^OPDRACHT$
				       ^Ba- ^Ma- ^MINOR- ^AR- ) ], $line );

      $self-> checkPositiveIntegerFields( [ qw( KP ) ], $line );

      $line->{CURRICULUM} ||= 0;
      $line->{ROOSTER} ||= 0;
      $line->{OPDRACHT} ||= 0;

      $self->checkEnumFields( [ qw( ^SP- ^VT- ^VP- ) ], 
			      [ '', qw( 1 2 3 4 12 34 ) ], $line );

      dieOnCurrentLine( "invalid line type " .
			"(must be one of '$self->{header}->{CURRICULUM}', " .
			"$self->{header}->{ROOSTER}' or '$self->{header}->{OPDRACHT}')\n" )
	unless ( $line->{CURRICULUM} + $line->{ROOSTER} + $line->{OPDRACHT} > 0 );

      dieOnCurrentLine( "you cannot combine the line type '$self->{header}->{CURRICULUM}'" .
			" with '$self->{header}->{ROOSTER}'" )
	if ( $line->{CURRICULUM} == 1 and $line->{ROOSTER} == 1 );

      ###############################################
      # do specific sanity checks and build database
      if ( $line->{CURRICULUM} ) {
	# check additional mandatory fields on curriculum line
	$self->checkMandatoryFields( [ qw( PROG DEELTIJDS TITULARIS DOCENT ) ], $line,
				     " on curriculum line" );
	$self->checkOptionalFieldsWithDefault( [ qw( CAT ) ], $line, '' );
	$self->checkOptionalFieldsWithDefault( [ qw( HC PR WC CP  ) ], $line, 0 );

	# check for fields that better be abscent on curriculum line
	$self->checkSuperfluousFields( [ qw( ACT GRP AGR ASS DSS TOTAAL
					     CAMPUS LOKAAL NEW ) ], $line,
				       " on curriculum line" );

	# check if we've got a decent value for deeltijds
	$self->checkEnumFields( [ qw( ^DEELTIJDS$ ) ], [ qw( 1 2 ) ], $line );

	# check if we're dealing with a proper program
	$self->checkEnumFields( [ qw( ^PROG$ ) ], [ qw( UA KdG Art ) ], $line );

	# check the appropriate number of SP
	$self->checkPositiveIntegerFields( [ qw( SP ) ], $line );

	# check if DOCENT is known
	if ( defined $pers ) {
	  $self->warnOnCurrentLine( "docent '$line->{DOCENT}' of curriculum line not found in personnel data\n" )
	    unless exists $pers->db()->{$line->{DOCENT}};
	}

	my $data = {
		    PROG             => $line->{PROG},
		    PROGJAAR         => $line->{PROGJAAR},
		    KP               => $line->{KP},
		    DEELTIJDS        => $line->{DEELTIJDS},
		    CAT              => $line->{CAT},
		    SEM              => $line->{SEM},
		    SP               => $line->{SP},
		    HC               => $line->{HC},
		    PR               => $line->{PR},
		    WC               => $line->{WC},
		    CP               => $line->{CP},
		    WP               => $line->{WP},
		    BP               => $line->{BP},
		    ST               => $line->{ST},
		    MP               => $line->{MP},
		    DOCENT           => $line->{DOCENT},
		    TIT              => $line->{TITULARIS},
		    OPM              => $line->{OPM}
		   };
	foreach my $key ( grep { m/^BA-|^MA-|^MINOR-|^AR-|^SP-|^VT-|^VP-/ } keys %$line ) {
	  $data->{$key} = $line->{$key};
	}
	$self->{data}->{$line->{SGNR}}->{C}->{$line->{OO}} = $data;

      } else {
	# all non curriculum lines should carry an activity and a (default) groupdescriptor
	$self->checkMandatoryFields( [ qw( ACT ) ], $line, " on non-curriculum line" );
	$self->checkOptionalFieldsWithDefault( [ qw( GRP ) ], $line,  '-' );

	# check for fields that better be abscent on non-curriculum lines
	$self->checkSuperfluousFields( [ qw( DEELTIJDS ) ], $line, 
				       " on non-curriculum line" );

	# check if we don't combine contact hour entries
	my $count = 0;
	foreach my $cu ( qw( HC PR WC CP WP BP ST MP ) ) {
	  ++$count if( $line->{$cu} );
	}
	$self->dieOnLine( "You cannot combine different activity types " .
			  "(Hoorcollege, Practicum, ..., Masterproef) " .
			  "on a single non-curriculum line. Please, split them.\n" )
	  if ( $count > 1 );
	$self->dieOnLine( "You cannot have no activity (Hoorcollege, Practicum, " .
			  "..., Masterproef) at all on a non-curriculum line. ".
			  "Please, complete the line.\n" )
	  if ( $count < 1 );
      }

      if ( $line->{ROOSTER} ) {
	# check if a room is designated
	$self-> checkMandatoryFields( [ qw( LOKAAL )], $line, " on rooster line" );

	# check if the campus is conformant and present
	$self->checkEnumFields( [ qw( ^CAMPUS$ ) ],
				[ qw( CPM CHO CGB CDD CST SCVO XXX CHO/CPM CHO/CST CPM/CST ) ],
				$line );


	# number of sessions and number of groups must be positive
	$self->checkPositiveIntegerFields( [ qw( ASS ) ], $line, 0 );
	$self->checkPositiveIntegerFields( [ qw( AGR ) ], $line, 1 );

	# session duration and total duration must be real and positive
	$self->checkPositiveRealFields( [ qw( DSS TOTAAL ) ], $line );

	$self->dieOnLine( "'$self->{header}->{TOTAAL}'-value is not correct!\n" )
	  if ( abs( $line->{ASS} * $line->{DSS} - $line->{TOTAAL} ) > 0.01 );

	$self->dieOnLine( "Duplicate line of type 'roosterlijn'\n" )
	  if ( exists( $self->{data}->{$line->{SGNR}}->{R}
		       ->{$line->{OO}}->{$line->{ACT}}->{$line->{GRP}} ) );

	my $hash =  {
		     SEM    => $line->{SEM},
		     DOCENT => $line->{DOCENT},
		     AGR    => $line->{AGR},
		     ASS    => $line->{ASS},
		     DSS    => $line->{DSS},
		     TOTAAL => $line->{TOTAAL},
		     CAMPUS => $line->{CAMPUS},
		     LOKAAL => $line->{LOKAAL},
		     OPM    => $line->{OPM}
		    };

	my $cu = 0;
	for my $tag ( qw( HC PR WC CP WP BP ) ) {
	  my $hours = $line->{$tag} if ( $line->{$tag} =~ /\d+.?\d*/ );
	  if ( $hours ) {
	    $self->dieOnLine( "More than one course-unit type on a single roster line " .
			      "is not allowed\n" )
	      if ( $cu );
	    $cu += $hours;
	    $hash->{CUTYPE} = $tag;
	  }
	}

	$hash->{CU} = $cu;

	$self->{data}->{$line->{SGNR}}->{R}
	  ->{$line->{OO}}->{$line->{ACT}}->{$line->{GRP}} = $hash;
      }

      if ( $line->{OPDRACHT} ) {
	# check for mandatory fields on 'opdracht' lines
	$self->checkMandatoryFields( [ qw( DOCENT )], $line, " on opdracht line" );

	# check if DOCENT is known
	if ( defined $pers ) {
	  $self->dieOnCurrentLine( "docent '$line->{DOCENT}' of opdracht line not found in personnel data\n" )
	    unless exists $pers->db()->{$line->{DOCENT}};
	}

	# check if the campus is conformant (but optional)
	$self->checkEnumFields( [ qw( ^CAMPUS$ ) ],
				[ qw( CPM CHO CGB CDD CST SCVO XXX ) ], 
				$line, 'optional' );

	# check positive integer fields
	$self->checkPositiveIntegerFields( [ qw( AGR ) ], $line, 1 );
	$self->checkPositiveIntegerFields( [ qw( GGR ) ], $line, 0 );

	# check if opdracht is a new one
	$self->checkBooleanFields( [ qw( ^NEW$ ) ], $line );
	$line->{NEW} ||= 0;

	my $hash =
	  $self
	    ->{data}->{$line->{SGNR}}->{O}->{$line->{OO}}
	      ->{$line->{DOCENT}}->{$line->{ACT}}->{$line->{GRP}} = {};

	foreach my $field ( qw( PROG PROGJAAR SEM GGR AGR CAMPUS NEW ) ) {
	  $hash->{$field} = $line->{$field};
	}

	foreach my $cu ( qw( HC PR WC CP WP BP ) ) {
	  $hash->{$cu} = $line->{$cu} if( $line->{$cu} );
	}

	foreach my $countable ( qw( ST MP ) ) {
	  $hash->{$countable} = $line->{$countable}
	    if ( $line->{$countable} );
	}

      }
    }
  }
  $self->_removeRaw();

  return 1;
}

1;


__END__

=head1 SEE ALSO

 --

=head1 COPYRIGHT

 CONFIDENTIAL AND PROPRIETARY (C) 2013 Digital Manifold Waves

=head1 AUTHOR

 Digital Manifold Waves -- F<walter.daems@ua.ac.be>

=cut

