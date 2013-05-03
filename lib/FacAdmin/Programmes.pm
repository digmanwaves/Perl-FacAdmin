# -*- cperl -*-
package FacAdmin::Programmes;

=head1 NAME

FacAdmin::Programmes -- Programmes database based on XLDB

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
  $self->_readHeader( 'Type',
		      {
		       '^Type$'                => 'TYPE',
		       '^Afkorting$'           => 'ACRONYM',
		       '^Omschrijving$'        => 'FULLNAME',
		       '^Minors$'              => 'MINORS',
		       '^Keuzepakketten$'      => 'KP',
		       '^Afstudeerrichtingen$' => 'AR',
		       '^Opmerking$'           => 'COMMENT',
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

      ###################
      # do sanity checks
      # check presence of mandatory fieldsheader
      $self->checkMandatoryFields( [ qw( TYPE ACRONYM FULLNAME ) ],
				   $line, "");
      $self->checkOptionalFieldsWithDefault( [ qw( MINORS KP AR COMMENT ) ],
					     $line, '' );

      $line->{TYPE} = uc( $line->{TYPE} );
      $self->{data}
	->{"$line->{TYPE}-$line->{ACRONYM}"} = {
						TYPE      => $line->{TYPE},
						ACRONYM   => $line->{ACRONYM},
						FULLNAME  => $line->{FULLNAME},
						MINORS    => $line->{MINORS},
						KP        => $line->{KP},
						AR        => $line->{AR},
						COMMENT   => $line->{COMMENT}
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

 CONFIDENTIAL AND PROPRIETARY (C) 2013 Walter Daems / Digital Manifold Waves

=head1 AUTHOR

 Digital Manifold Waves -- F<walter@digmanwaves.net>

=cut
