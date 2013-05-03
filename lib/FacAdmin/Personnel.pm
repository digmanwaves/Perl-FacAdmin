# -*- cperl -*-
package FacAdmin::Personnel;

=head1 NAME

FacAdmin::Personnel -- Personnel database based on XLDB

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

  ConLogger::logitem( "Reading header" );
  $self->_readHeader( 'Achternaam',
		      {
		       '^Achternaam$'      => 'LASTNAME',
		       '^Voornaam$'        => 'FIRSTNAME',
		       '^Naam$'            => 'NAME',
		       '^Email$'           => 'EMAIL',
		       '^Opleiding$'       => 'OPL',
		       '^Opdracht$'        => 'VTE',
		       '^Onderzoek$'       => 'OZP',
		       '^Dienstverlening$' => 'DVP',
		      }
		    );

  ConLogger::logitem( "Reading data records" );

  {
    my $pb = ConLogger::ProgressBar->new();

    for ( my $i = $self->property('HeaderRow')+1; $i <= $self->property('MaxRow'); ++$i ) {

      $pb->progress( $i, $self->property('MaxRow') );

      ######################################
      # parse the line into the %$line hash
      my $line = $self->parseLine( $i );

      ###################
      # do sanity checks
      # check presence of mandatory fields
      $self->checkMandatoryFields( [ qw( NAME OPL VTE OZP DVP ) ],
				   $line, "" );

      $self->checkEnumFields( [ qw( ^OPL$ ) ],
			      [ qw( BK CH EM EI AV PO TEW ) ],
			      $line );

      $self->checkPercentageFields( [ qw( VTE OZP DVP ) ],
				    $line );

      $self->{data}->{$line->{NAME}} = {
					LASTNAME  => $line->{LASTNAME},
					FIRSTNAME => $line->{FIRSTNAME},
					EMAIL     => $line->{EMAIL},
					OPL       => $line->{OPL},
					VTE       => $line->{VTE},
					OZP       => $line->{OZP},
					DVP       => $line->{DVP},
				       };
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

