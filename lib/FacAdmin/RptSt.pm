# -*- cperl -*-
package FacAdmin::RptSt;

=head1 NAME

FacAdmin::RptSt -- Report Study Guide

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
  $self->_readHeader( 'Studiegidsnr.',
		      {
		       '^Studiegidsnr.$'                    => 'SGNR',
		       '^Omschrijving$'                     => 'SHORTDESCR-NL',
		       '^Vert. omsch.$'                     => 'SHORTDESCR-EN',
		       '^Lange naam studiedeel$'            => 'FULLDESCR-NL',
		       '^Vert. lange naam studiedeel'       => 'FULLDESCR-EN',
		       '^Dipl. omschr.$'                    => 'DIPL-NL',
		       '^Vert. Dipl.$'                      => 'DIPL-NL',
		       '^Organisatie$'                      => 'ORG',
		       '^Loopbaan$'                         => 'PROG',
		       '^Vakgebied$'                        => 'DOMAIN',
		       '^Fiat$'                             => 'FIAT',
		       '^Studiedeel vrijgeven$'             => 'RELEASED',
		       '^In studiegids$'                    => 'PUBLISHED',
		       '^Doorsneeaanbod$'                   => 'PROGYEAR',
		       '^Eenheden$'                         => 'SP',
		       '^Contacturen$'                      => 'CU',
                       '^Campus$'                           => 'CAMPUS',
		       '^Vereistengroep$'                   => 'REQGROUP',
		       '^Studiedeelkenmerken$'              => 'PROPS',
		       '^Onderdeel: Primair en Beoordeeld$' => 'PARTS',
		       '^Equivalent OO$'                    => 'EQUIVALENCE',
		      }
		    );
  $self->_registerValueHeaderEntry( 'PROPS', 'PROPS-value' );
  $self->_registerValueHeaderEntry( 'PARTS', 'PARTS-1' );
  $self->_registerValueHeaderEntry( 'PARTS-1', 'PARTS-2' );

  ConLogger::logitem( "Reading data records" );

  {
    my $pbar = ConLogger::ProgressBar->new();

    for ( my $i = $self->property('HeaderRow')+1; $i <= $self->property('MaxRow'); ++$i ) {

      $pbar->progress( $i, $self->property('MaxRow') );

      ######################################
      # parse the line into the %$line hash

      my $line = $self->parseLine( $i );

      ##############
      # process first line (OO-line)
      my $sgnr = $line->{SGNR};
      $self->{data}->{$sgnr} = {};

      print STDERR "LINE $i : " . $sgnr . "\n";

      # read atomic data
      for my $label ( qw( SHORTDESCR-NL SHORTDESCR-EN FULLDESCR-NL FULLDESCR-EN
      			  DIPL-NL DIPL-NL ORG PROG DOMAIN FIAT RELEASED PUBLISHED
      			  PROGYEAR SP CU CAMPUS REQGROUP EQUIVALENCE ) ) {
      	$self->{data}->{$sgnr}->{$label} = $line->{$label};
      }

      $self->{data}->{$sgnr}->{PROPS} = 
	{ $line->{PROPS} => [ $line->{'PROPS-value'} ] };
      $self->{data}->{$sgnr}->{PARTS} = 
	{ $line->{PARTS} => $line->{'PARTS-1'} . '-' . $line->{'PARTS-2'} };

      # read other lines until new OO-line starts
      for( ++$i; $i <= $self->property('MaxRow'); ++$i ) {
	$line = $self->parseLine( $i );

	if ( defined $line->{SGNR} ) {
	  # this is the first line of a new OO
	  --$i;
	  print STDERR "rewinding LINE $i\n";
	  last;
	}
	else {
	  print STDERR "dealing continuation LINE $i\n";
	  # this is a continuation line
	  if ( exists $line->{PROPS} ) {
	    if ( exists $self->{data}->{$sgnr}->{PROPS}->{$line->{PROPS}} ) {
	      push @{$self->{data}->{$sgnr}->{PROPS}->{$line->{PROPS}}},
		$line->{'PROPS-value'};
	    }
	    else {
	      $self->{data}->{$sgnr}->{PROPS}->{$line->{PROPS}} = [ $line->{'PROPS-value'} ];
	    }
	  }

	  if ( exists $line->{PARTS} ) {
	    $self->{data}->{$sgnr}->{PARTS}->{$line->{PARTS}} = 
	      $line->{'PARTS-1'} . '-' . $line->{'PARTS-2'};
	  }
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

 CONFIDENTIAL AND PROPRIETARY (C) 2013 Walter Daems / Digital Manifold Waves

=head1 AUTHOR

 Digital Manifold Waves -- F<walter@digmanwaves.net>

=cut

