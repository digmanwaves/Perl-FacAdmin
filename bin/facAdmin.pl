#!/usr/bin/perl -w
# -*- cperl -*-

use strict;
use 5.010;

use Getopt::Long qw(:config no_auto_abbrev);
use Pod::Usage;

use IO::File;
use File::Basename;
use File::Spec;

use ConLogger;
use ConLogger::SubTask;
use XLDB::IFile;
use FacAdmin::Personnel;
use FacAdmin::Programmes;
use FacAdmin::EduInventory;
use FacAdmin::RptSt;

use XLDB::OFile;
use XLDB::Sheet;

use Excel::Writer::XLSX;	# ppm install Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;

use Encode;

use Data::Dumper;
$Data::Dumper::Sortkeys = 1;

END { print "\n"; }

sub help;

sub checkHeaderEntry;

my $toolname = 'facAdmin';
my $company  = 'Digital Manifold Waves';
my $author   = 'Walter Daems (walter.daems@ua.ac.be)';
my $date     = '2013/04/09';
my $version  = '1.00';

print
  "/***************************************************\n"
  . " * $company - $toolname\n"
  . " * Author : $author\n"
  . " * Version: $version ($date)\n"
  . " ***************************************************/\n";

my ($program, $installdir) = fileparse( $0 );

my $tasks;
my $icheck;
my $ocheck;
my $curriculum;
my $roster;
my $help;
my $man;
my $limit;

GetOptions( "limit=s"         => \$limit,
	    "tasks"           => \$tasks,
	    "internal-check"  => \$icheck,
	    "sisa-check"      => \$ocheck,
	    "curriculum"      => \$curriculum,
	    "roster"          => \$roster,
	    "help|?"          => \$help,
	    "man"             => \$man )
  or pod2usage(2);

pod2usage(1) if $help;
pod2usage(-exitstatus => 0, -verbose => 2 ) if $man;
pod2usage(1) unless ( defined( $icheck ) or
		      defined( $ocheck ) or
		      defined( $tasks ) or
		      defined( $curriculum ) or
		      defined( $roster ) );

my @iFileNames = @ARGV;
my ( undef, $path, $suffix ) = fileparse($iFileNames[0], qw( .xlsx .xlsm ) );

####################
# Parse input files
##################


ConLogger::logitem( 'Scanning input files' );
my $sheets = {
	      'Personeelsoverzicht' => [],
	      'Programmaoverzicht' => [],
	      'Onderwijsoverzicht' => [],
	     };
{
  my $st = ConLogger::SubTask->new();
  foreach my $iFileName ( @iFileNames ) {
    ConLogger::logitem( $iFileName );
    my $xldbfile = XLDB::IFile->new();
    $xldbfile->open( $iFileName );
    while( my ( $type, $array ) = each %$sheets ) {
      my $sheet = $xldbfile->sheet( $type );
      if ( defined $sheet ) {
	my $sh_fname = [ $sheet, $iFileName ];
	push @$array, $sh_fname;
      }
    }
  }
}

ConLogger::logitem( 'Reading personnel data' );
my $pers = FacAdmin::Personnel->new();
{
  my $st = ConLogger::SubTask->new();
  foreach my $sh_fname ( @{$sheets->{Personeelsoverzicht}} ) {
    ConLogger::logitem( "Reading data from file '$sh_fname->[1]'" );
    my $st = ConLogger::SubTask->new();
    $pers->connect( $sh_fname->[0] );
    $pers->buildDataBase();
  }
}

ConLogger::logitem( 'Reading programme data' );
my $prog = FacAdmin::Programmes->new();
{
  my $st = ConLogger::SubTask->new();
  foreach my $sh_fname ( @{$sheets->{Programmaoverzicht}} ) {
    ConLogger::logitem( "Reading data from file '$sh_fname->[1]'" );
    my $st = ConLogger::SubTask->new();
    $prog->connect( $sh_fname->[0] );
    $prog->buildDataBase();
  }
}

ConLogger::logitem( 'Reading educational inventory' );
my $einv = FacAdmin::EduInventory->new();
{
  my $st = ConLogger::SubTask->new();
  foreach my $sh_fname ( @{$sheets->{Onderwijsoverzicht}} ) {
    ConLogger::logitem( "Reading data from file '$sh_fname->[1]'" );
    my $st = ConLogger::SubTask->new();
    $einv->connect( $sh_fname->[0] );
    $einv->buildDataBase( $pers );
  }
}

#######################
# sanity checks on owo
#####################

if ( $icheck ) {
  ConLogger::logitem( 'Performing relational sanity checks' );

  while ( my ( $sgnr, $CRO ) = each( %{$einv->db()} ) ) {
    # check duplicate use of Studiegidsnrs
    foreach my $cro ( qw( C R O ) ) {
      my @oos = keys %{$CRO->{$cro}};
      if ( @oos > 1 ) {
	die( "Error: Duplicate use of " . labelof( $einv->header()->{SGNR} ) .
	     " '$sgnr'\n" .
	     "       You used it for: " . join( ", ", @oos ) . "\n" );
      }
    }
  }
  exit(0);
}

#######################################
# crosschecking with QQ_RPTST van SisA
#####################################

if ( $ocheck ) {
  ConLogger::logitem( 'Performing cross-check with SisA' );

  while ( my ( $sgnr, $CRO ) = each( %{$einv->db()} ) ) {
    # check duplicate use of Studiegidsnrs
    foreach my $cro ( qw( C R O ) ) {
      my @oos = keys %{$CRO->{$cro}};
      if ( @oos > 1 ) {
	die( "Error: Duplicate use of " . labelof( $einv->header()->{SGNR} ) .
	     " '$sgnr'\n" .
	     "       You used it for: " . join( ", ", @oos ) . "\n" );
      }
    }
  }
  exit(0);
}

my ( $sec, $min, $hour, $mday, $mon, $year ) = localtime time;
$year += 1900; ++$mon;

#########################
# generate task database
#######################

if ( $tasks ) {
  ConLogger::logitem( 'Generating task database' );

  my $taskdb = {};
  {
    my $pb = ConLogger::ProgressBar->new();
    my $sgnrcount = keys %{$einv->db()};
    my $counter = 0;

    while ( my ( $sgnr, $cro ) = each( %{$einv->db()} ) ) {
      $pb->progress( $counter++, $sgnrcount );

      my $o = $cro->{O};
    OO:
      while ( my ( $oo, $docenten ) = each( %$o ) ) {
	# loop will only execute once because of check of duplicate SGNRs
	while ( my ( $doc, $activiteiten ) = each( %$docenten ) ) {
	  while ( my ( $act, $groepen ) = each( %$activiteiten ) ) {

	    # find active tag
	    my $tag;
	    my $sem;
	    my $progjaar;
	    my $newact;
	    my $tagvalue;
	    my $groupcount = 0;
	    my $groupoversize = 0;

	    while ( my ( $grp, $detail ) = each ( %$groepen ) ) {

	      # temporary to avoid 3de jaar nieuw curriculum
	      next OO if ( $detail->{PROGJAAR} == 3
			   and
			   $detail->{PROG} eq "UA" );

	      # find active tag
	      if ( defined( $tag ) ) {
		# to avoid string comparisons in ==
		# = temporary hack
		if ( $tag ne "PROG" ) {
		
		  # check if new tags corresponds to existing ones
		  die( "Error: Docent '$doc' has an inconsistent workload " .
		       "for $oo - $act for the different groups (s)he's teaching. " .
		       "Correct this in the sheet 'Onderwijsoverzicht'.\n" )
		    unless ( $tagvalue == $detail->{$tag} );
		  die( "error: Docent '$doc' has inconsistent semesters " .
		       "for $oo - $act for the different groups (s)he's teaching. " .
		       "Correct this in the sheet 'Onderwijsoverzicht'.\n" )
		    unless ( $sem == $detail->{SEM} );
		  die( "error: Docent '$doc' has inconsistent program year " .
		       "for $oo - $act for the different groups (s)he's teaching. " .
		       "Correct this in the sheet 'Onderwijsoverzicht'.\n" )
		    unless ( $progjaar == $detail->{PROGJAAR} );
		
		  die( "error: Docent '$doc' has inconsistent 'nieuwe activiteit'-marker " .
		       "for $oo - $act for the different groups (s)he's teaching. " .
		       "Correct this in the sheet 'Onderwijsoverzicht'.\n" )
		    unless ( $newact == $detail->{NEW} );
		}
	
	      } else {
		# find first tag
		$tag = findActiveCUTag( $detail ) unless( defined( $tag ) );
		$tagvalue = $detail->{$tag};
		$sem      = $detail->{SEM};
		$progjaar = $detail->{PROGJAAR};
		$newact   = $detail->{NEW};
	      }

	      $groupcount    += $detail->{AGR};
	      $groupoversize += $detail->{GGR};
	    }

	    $newact = 0 unless defined( $newact );
	    $taskdb->{$doc}->{$sem}->{$progjaar}->{"$oo - $act"} =
	      {
	       $tag => $tagvalue,
	       AGR  => $groupcount,
	       GGR  => $groupoversize,
	       NEW  => $newact,
	      };
	  }
	}
      }
    }
  }

  #########################
  # generate taskoverviews
  #######################

  ConLogger::logitem( "Generating tasks in file 'Opdrachten.xlsx'" );

  # Create a new Excel workbook
  my $oFileName = File::Spec->catfile( $path, "Opdrachten.xlsx" );
  my $opdrachtenbook =
    XLDB::OFile->new( filename => $oFileName,
		      title    => 'FacAdmin - Opdrachten',
		      author   => 'Walter Daems',
		      manager  => 'Walter Daems',
		      company  => 'Universiteit Antwerpen',
		      division => 'Faculteit Toegepaste Ingenieurswetenschappen',
		      toolname => 'Gegenereerd met Digital Manifold Waves - FacAdmin' );

  ####################
  # Setup index sheet
  ##################

  my $idxsh = $opdrachtenbook->makeSheet( 'Index', 'P' );

  my $idxrow = 0;
  $idxsh->write( $idxrow, 0, 'Opdrachtfiche-index', 't' );
  $idxsh->write( ++$idxrow, 0,
		 'Klik op de namen hieronder om naar de corresponderende opdrachtfiche ' .
		 'te gaan.', 'LW' );

  $idxrow +=1;
  my $idxcol = 0;
  $idxsh->write( ++$idxrow, $idxcol, 'Naam', 'Lu' );
  $idxsh->write( $idxrow, ++$idxcol, 'Opleiding', 'Lu' );
  $idxsh->set_column( $idxcol, $idxcol, 12, 'C' );
  $idxsh->write( $idxrow, ++$idxcol, 'Aanstelling', 'Lu' );
  $idxsh->set_column( $idxcol, $idxcol, 12, 'RI1P1' );
  $idxsh->write( $idxrow, ++$idxcol, 'Onderwijs', 'Lu' );
  $idxsh->set_column( $idxcol, $idxcol, 12, 'RI1P1' );
  $idxsh->write( $idxrow, ++$idxcol, 'Onderzoek', 'Lu' );
  $idxsh->set_column( $idxcol, $idxcol, 12, 'RI1P1' );
  $idxsh->write( $idxrow, ++$idxcol, 'Dienstverlening', 'Lu' );
  $idxsh->set_column( $idxcol, $idxcol, 12, 'RI1P1' );
  $idxsh->write( $idxrow, ++$idxcol, 'Marge', 'Lu' );
  $idxsh->set_column( $idxcol, $idxcol, 12, 'RI1P1' );

  my $idxfirstrow = $idxrow;
  $idxsh->set_column( 0, 0, 60 );


  #########################
  # Write opdrachtenfiches
  #######################

  {
    my $pb = ConLogger::ProgressBar->new();
    my $docentcount = keys %$taskdb;
    my $docentctr = 0;

    for my $doc ( sort keys %$taskdb ) {
      $pb->progress( ++$docentctr, $docentcount );

      die( "Error: could not find docent '$doc' in Sheet 'Personeelsoverzicht'. " .
	   "Please, complete the personnel data.\n" )
	unless ( defined $pers and defined $pers->db() and exists $pers->db()->{$doc} );

      if ( defined( $limit ) ) {
	next unless $pers->db()->{$doc}->{OPL} =~ /$limit/;
      }

      my $sem = $taskdb->{$doc};

      my $mydoc = $doc;
      $mydoc =~ s/'//g;	# for Bob T'Jollyn
      my $wsh = $opdrachtenbook->makeSheet( $mydoc, 'L' );

      my $leftcol = 0;
      my $rightcol = 22;

      $wsh->insert_image( 0, $rightcol - 6, 'UALogo.png' );

      # Write title on sheet
      my ($col , $row) = ( 0, 0);
      $wsh->write( $row, $col, "Opdrachtenfiche van $doc", 't' );
      ++$row;
      ++$row;
      $wsh->merge_range( $row, $leftcol, $row, $rightcol - 8,
			 "Deze fiche beschrijft de minimale verwachtingen van de faculteit " .
			 "op het vlak van je tijdsbesteding voor de drie kerntaken van de " .
			 "universiteit. Het verschil tussen je effectieve inzet (die zelfs " .
			 "meer dan 100% kan bedragen) is vrije ruimte ('marge') die je naar " .
			 "eigen inzicht kan invullen in elk van de drie kerntaken (onderwijs, " .
			 "onderzoek en dienstverlening).", 'LW' );
      $wsh->set_row( $row, 22 );

      ++$row;
      $wsh->merge_range( $row, $leftcol, $row, $rightcol - 8,
			 "De faculteit engageert zich om in in overleg en in functie van je " .
			 "statuut een voldoende groot percentage onderzoekstijd vrij te maken " .
			 "zolang een aanvaardbare onderzoeksoutput meetbaar is.", 'LW' );
      $wsh->set_row( $row, 12 );

      ++$row;
      $wsh->merge_range( $row, $leftcol, $row, $rightcol - 8,
			 "De faculteit streeft evenwicht na tussen alle medewerkers in gelijke " .
			 "categorieën, maar heeft ook het recht om het geheel van onderwijs- " .
			 "en dienstverleningstaken aan te vullen tot de resterende marge nul " .
			 "bedraagt.", 'LW' );
      $wsh->set_row( $row, 22 );

      ++$row;
      $wsh->merge_range( $row, $leftcol, $row, $rightcol - 8,
			 "Opmerking: markeer wijzigingen, noteer de datum en je naam " .
			 "zodat de administratie je wijzigingen kan aanbrengen " .
			 "in de database.", 'LW' );
      $wsh->set_row( $row, 12 );


      $row += 2;
      $col = 0;
      $wsh->write( $row,   $col++, "Opdrachtspercentage:", 'LxP2' );
      $wsh->merge_range( $row, $col,   $row, $col+1, $pers->db()->{$doc}->{VTE}, 'LxP2' );
      $wsh->merge_range( $row, $col+2, $row, $rightcol, undef, 'LxP2' );
      my ( $oprow, $opcol ) = ( $row, $col );

      $row += 2;
      $col = 0;
      $wsh->write( $row,   $col++, "Onderwijs:", 'LgP2' );
      $wsh->merge_range( $row, $col,   $row, $col+1, 0.4, 'LgP2' );
      $wsh->merge_range( $row, $col+2, $row, $rightcol, undef, 'LgP2' );
      my ( $owrow, $owcol ) = ( $row, $col );

      ######################
      # Generate owo header

      $col = $leftcol;
      $row += 5;
      $wsh->write( $row,   $col, "OO - Activiteit", 'Lu' );
      $wsh->set_column( $col, $col, 60 );
      $wsh->write( $row, ++$col, "Programmajaar", 'Tu' );
      $wsh->set_column( $col, $col, 6 );

      my $totcol = ++$col;
      my $sumcol = $totcol;
      $wsh->write( $row,   $col, "HC [h]", 'Tu' );
      $wsh->write( $row, ++$col, "WC/Seminarie [h]", 'Tu' );
      $wsh->write( $row, ++$col, "PR [h]", 'Tu' );
      $wsh->write( $row, ++$col, "Coaching [h]", 'Tu' );
      $wsh->write( $row, ++$col, "Wetenschappelijk project [h]", 'Tu' );
      $wsh->write( $row, ++$col, "Bachelorproef [h]", 'Tu' );
      $wsh->set_column( $totcol, $col, 7 );
      $wsh->merge_range( $row-1, $totcol, $row-1, $col, "Totaal", 'Ca' );
      my $endtotcol = $col;

      $wsh->write( $row, ++$col, "Aantal groepen [-]", 'Tu' );
      $wsh->set_column( $col, $col, 4 );
      my $agrcol = $col;

      my $detcol = ++$col;
      $wsh->write( $row,   $col, "HC [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "WC/Seminarie [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "PR [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "HC-parallel [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "WC-parallel [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "PR-parallel [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "Coaching [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "Wetensch. project [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "Bachelorproef [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "Toeslag nieuw OO [h/wk/sem]", 'Tu' );
      $wsh->write( $row, ++$col, "Toeslag gr. groep [aantal]", 'Tu' );
      my $middetcol = $col;
      $wsh->write( $row, ++$col, "Stage [aantal]", 'Tu' );
      $wsh->write( $row, ++$col, "Masterproef [aantal]", 'Tu' );
      $wsh->write( $row, ++$col, "Andere taken [VTE]", 'Tu' );
      $wsh->set_column( $detcol, $col, 7 );
      $wsh->merge_range( $row-1, $detcol, $row-1, $col, "Detail", 'Ca' );
      my $enddetcol = $col;
      my $endsumcol = $col;

      $row -=2;
      $col = $detcol;
      my $coeffrow = $row;
      $wsh->write( $row,   $col, 3.0 / 100, 'RI1P1' ); # HC
      $wsh->write( $row, ++$col, 3.0 / 100, 'RI1P1' ); # WC
      $wsh->write( $row, ++$col, 3.0 / 100, 'RI1P1' ); # PR
      $wsh->write( $row, ++$col, 2.0 / 100, 'RI1P1' ); # HC parallel
      $wsh->write( $row, ++$col, 2.0 / 100, 'RI1P1' ); # WC parallel
      $wsh->write( $row, ++$col, 2.0 / 100, 'RI1P1' ); # PR parallel
      $wsh->write( $row, ++$col, 1.0 / 100, 'RI1P1' ); # Coaching
      $wsh->write( $row, ++$col, 1.0 / 100, 'RI1P1' ); # Wetenschappelijke project
      $wsh->write( $row, ++$col, 1.3 / 100, 'RI1P1' ); # Bachelorproef
      $wsh->write( $row, ++$col, 1.0 / 100, 'RI1P1' ); # toeslag nieuw
      $wsh->write( $row, ++$col, 1.0 / 100, 'RI1P1' ); # toeslag grote groep
      $wsh->write( $row, ++$col, 0.6 / 100, 'RI1P1' ); # Stage
      $wsh->write( $row, ++$col, 2.0 / 100, 'RI1P1' ); # Masterproef
      $wsh->write( $row, ++$col, 5.0 / 100, 'RI1P1' ); # Andere taken
      $wsh->merge_range( $row-1, $detcol, $row-1, $col,
			 "Opdrachtscoëfficiënten [%/eenheid]", 'Ca' );
      my $firstdatarow = $row += 3;

      ###########################
      # generate owo data lines

      my $firststartrow;
      foreach my $semester ( sort keys %$sem ) {
	my $progjaren = $sem->{$semester};
	$wsh->set_row( $row, 5 );
	$wsh->write( ++$row, 0, "Semester $semester:", 'Lb' );
	++$row;
	my $startrow = $row;
	$firststartrow ||= $row;

	foreach my $progjaar ( sort keys %$progjaren ) {
	  my $ooacts = $progjaren->{$progjaar};
	  foreach my $ooact ( sort keys %$ooacts ) {
	    my $detail = $ooacts->{$ooact};
	    $col = $leftcol;
	    my $grgractcount = 0;
	    my $newacthours = 0;
	    $wsh->write( $row,   $col, $ooact, 'LI2' );
	    $wsh->write( $row, ++$col, $progjaar, 'C' );
	    foreach my $tag ( qw( HC WC PR ) ) {
	      ++$col;
	      if ( exists( $detail->{$tag} ) ) {
		$wsh->write( $row, $col, $detail->{$tag}, 'RI1F1' );
		my $primaryhours = $detail->{$tag} / 12;
		$grgractcount += 1;
		$newacthours += $primaryhours;
		my $secondaryhours = $primaryhours * ( $detail->{AGR} - 1 );
		$wsh->write( $row, $col + $detcol - $sumcol, $primaryhours, 'RI1F1' );
		$wsh->write( $row, $col + $detcol - $sumcol + 3, $secondaryhours, 'RI1F1' )
		  if ( $secondaryhours );
	      }
	    }
	    foreach my $tag ( qw( CP WP BP ) ) {
	      ++$col;
	      if ( exists( $detail->{$tag} ) ) {
		$wsh->write( $row, $col, $detail->{$tag}, 'RI1F1' );
		my $primaryhoursonly = $detail->{$tag} / 12;
		$newacthours += $primaryhoursonly * $detail->{AGR};
		$wsh->write( $row, $col + $detcol - $sumcol + 3, $newacthours, 'RI1F1' );
	      }
	    }
	    $wsh->write( $row, ++$col, $detail->{AGR}, 'RI1' );
	    $col += 9; # get us from the 'Aantal groepen' to before the 'Toeslag nieuw' column

	    ++$col;
	    my $newact = $newacthours * $detail->{NEW};
	    $wsh->write( $row, $col, $newact, 'RI1F1' ) if $newact;

	    ++$col;
	    my $grgr = $grgractcount * $detail->{GGR};
	    $wsh->write( $row, $col, $grgr, 'RI1' ) if $grgr;

	    foreach my $tag ( qw( ST MP ) ) {
	      ++$col;
	      if ( exists( $detail->{$tag} ) ) {
		$wsh->write( $row, $agrcol, $detail->{$tag}, 'RI1' );
		$wsh->write( $row, $col, $detail->{$tag}, 'RI1' );
		my $primaryhoursonly = $detail->{$tag};
		
	      }
	    }
	    ++$row;
	  }
	  $col = $leftcol;
	}
	$wsh->write( $row, $col++, "Sem-$semester - subtotalen:", 'Rb' );
	for ( $col = $totcol; $col <= $endtotcol; ++$col ) {
	  $wsh->write_formula( $row,   $col,
			       "=SUBTOTAL(9," .
			       xl_range( $startrow, $row-1, $col, $col ) . ")",
			       'RI1bF1' );
	}
	for ( $col = $detcol; $col < $middetcol; ++$col ) {
	  $wsh->write_formula( $row,   $col,
			       "=SUBTOTAL(9," .
			       xl_range( $startrow, $row-1, $col, $col ) . ")",
			       'RI1bF1' );
	}
	for ( $col = $middetcol; $col < $enddetcol; ++$col ) {
	  $wsh->write_formula( $row,   $col,
			       "=SUBTOTAL(9," .
			       xl_range( $startrow, $row-1, $col, $col ) . ")",
			       'RI1b' );
	}

	$col = $leftcol;
	++$row;
	$wsh->write( $row, $col, "Gemiddeld #CU/week:", 'Rb' );
	++$col;

	if ( $startrow == $row - 2 ) {
	  $wsh->write_formula( $row, $col,
			       'sum(' . 
			       xl_range( $startrow, $startrow, $totcol, $endtotcol ) .
			       ')*' . xl_rowcol_to_cell( $startrow, $endtotcol+1) . '/12',
			       'RI1bF1' );
	}
	else {
	  my @sumcomponents;

	  for( my $i = $totcol; $i <= $endtotcol; ++$i ) {
	    push @sumcomponents, "sumproduct(" .
	      xl_range( $startrow, $row-2, $i, $i ) . "," .
		xl_range( $startrow, $row-2, $endtotcol+1, $endtotcol+1 ) . ")";
	  }
	  $wsh->write_formula( $row, $col,
			       '(' . join( '+', @sumcomponents ) . ')/12',
			       'RI1bF1' );
	}
	++$row;
	$wsh->set_row( $row, 5 );
	for ( my $i = $leftcol+1; $i <= $rightcol; ++$i ) {
	  $wsh->write( $row, $i, undef, 'Lu' );
	}
	++$row;
      }
      $wsh->write( $row-1, 0, undef, 'Lu' );

      $wsh->write( $row, $sumcol - 2, "Jaar-subtotalen:", 'Rb' );
      for ( $col = $totcol; $col <= $endtotcol; ++$col ) {
	$wsh->write_formula( $row,   $col,
			     "=SUBTOTAL(9," .
			     xl_range( $firstdatarow, $row-1, $col, $col ) . ")",
			     'RI1bF1' );
      }
      for ( $col = $detcol; $col < $middetcol; ++$col ) {
	$wsh->write_formula( $row,   $col,
			     "=SUBTOTAL(9," .
			     xl_range( $firstdatarow, $row-1, $col, $col ) . ")",
			     'RI1bF1' );
      }
      for ( $col = $middetcol; $col < $enddetcol; ++$col ) {
	$wsh->write_formula( $row,   $col,
			     "=SUBTOTAL(9," .
			     xl_range( $firstdatarow, $row-1, $col, $col ) . ")",
			     'RI1b' );
      }
      $wsh->write_formula( $row, $col, "=" . xl_rowcol_to_cell( $oprow, $opcol ), 'RI1bF1' );

      $wsh->write_formula( $owrow, $owcol, "=SUMPRODUCT(" .
			   xl_range( $coeffrow, $coeffrow, $detcol, $enddetcol ) . "," .
			   xl_range( $row, $row, $detcol, $enddetcol ) . ")",
			   'LgP2' );

      ++$row;
      $wsh->write( $row, $sumcol - 2, "#Gemiddeld #CU/week:", 'Rb' );


      if ( $firststartrow == $row - 5 ) {
	$wsh->write_formula( $row, $col,
			     'sum(' . 
			     xl_range( $firststartrow, $firststartrow, $totcol, $endtotcol ) .
			     ')*' . xl_rowcol_to_cell( $firststartrow, $endtotcol+1) . '/24',
			     'RI1bF1' );
      }
      else {
	my @sumcomponents;

	for( my $i = $totcol; $i <= $endtotcol; ++$i ) {
	  push @sumcomponents, "sumproduct(" .
	    xl_range( $firststartrow, $row-5, $i, $i ) . "," .
	      xl_range( $firststartrow, $row-5, $endtotcol+1, $endtotcol+1 ) . ")";
	}
	$wsh->write_formula( $row, $sumcol-1,
			     '(' . join( '+', @sumcomponents ) . ')/24',
			     'RI1bF1' );
      }

      $row += 2;
      $col = $leftcol;
      $wsh->write( $row, $col++, "Onderzoek:", 'LgP2' );
      $wsh->merge_range( $row, $col,   $row, $col+1, $pers->db()->{$doc}->{OZP}, 'LgP2' );
      $wsh->merge_range( $row, $col+2, $row, $rightcol, undef, 'LgP2' );
      my ( $ozrow, $ozcol ) = ( $row, $col );

      $row += 2;
      $col = $leftcol;
      $wsh->write( $row, $col++, "Dienstverlening:", 'LgP2' );
      $wsh->merge_range( $row, $col,   $row, $col+1, $pers->db()->{$doc}->{DVP}, 'LgP2' );
      $wsh->merge_range( $row, $col+2, $row, $rightcol, undef, 'LgP2' );
      my ( $dvrow, $dvcol ) = ( $row, $col );

      my $marginformula =
	sprintf( "=%s-%s-%s-%s",
		 xl_rowcol_to_cell( $oprow, $opcol ),
		 xl_rowcol_to_cell( $owrow, $owcol ),
		 xl_rowcol_to_cell( $ozrow, $ozcol ),
		 xl_rowcol_to_cell( $dvrow, $dvcol ) );


      $row += 2;
      $col = $leftcol;
      $wsh->write( $row, $col++, "Marge:", 'LxP2' );
      $wsh->merge_range( $row, $col,   $row, $col+1, $marginformula, 'LxP2' );
      $wsh->merge_range( $row, $col+2, $row, $rightcol, undef, 'LxP2' );
      my ( $mgrow, $mgcol ) = ( $row, $col );

      ######################
      # Complete indexsheet
      ####################

      $idxcol = 0;
      #my $mydoc = $doc;
      #$mydoc =~ s/'//g;	# for Bob T'Jollyn
      $idxsh->write_url( ++$idxrow, $idxcol,
			 "internal:'$mydoc'!A1", 'link', $doc );
      $idxsh->write( $idxrow, ++$idxcol, $pers->db()->{$doc}->{OPL} );
      $idxsh->write( $idxrow, ++$idxcol,
		     "='$mydoc'!" . xl_rowcol_to_cell( $oprow, $opcol ) );
      $idxsh->write( $idxrow, ++$idxcol,
		     "='$mydoc'!" . xl_rowcol_to_cell( $owrow, $owcol ) );
      $idxsh->write( $idxrow, ++$idxcol,
		     "='$mydoc'!" . xl_rowcol_to_cell( $ozrow, $ozcol ) );
      $idxsh->write( $idxrow, ++$idxcol,
		     "='$mydoc'!" . xl_rowcol_to_cell( $dvrow, $dvcol ) );
      $idxsh->write( $idxrow, ++$idxcol,
		     "='$mydoc'!" . xl_rowcol_to_cell( $mgrow, $mgcol ) );
    }
  }

  ######################
  # Finalize indexsheet
  ####################
  $idxsh->autofilter( $idxfirstrow, 0, $idxrow, $idxcol );

  $opdrachtenbook->close();
}


if ( $curriculum ) {

  ConLogger::logitem( 'Generating curriculum database' );
  my $currdb = {};
  {
    my $pb = ConLogger::ProgressBar->new();
    my $sgnrcount = keys %{$einv->db()};
    my $sgnrctr = 0;

    while ( my ( $sgnr, $cro ) = each( %{$einv->db()} ) ) {

      $pb->progress( ++$sgnrctr, $sgnrcount );

      my $c = $cro->{C};
      while ( my ( $oo, $details ) = each( %{$c} ) ) {
	# loop will only execute once because of check of duplicate SGNRs

	# geneate data structure (linear line data)
	my $data = { OO => $oo };
	foreach my $d ( qw( SP HC PR TIT DEELTIJDS ) ) {
	  $data->{$d} = $details->{$d};
	  $data->{$d} =~ s/-,//g;
	  $data->{$d} =~ s/, -//g;
	}

	$data->{I} = 0;
	foreach my $d ( qw( WC CP ) ) {
	  $data->{I} += $details->{$d};
	}

	# write linear line data into every applicable programme
	my @regopl = grep { m/^BA-|^MA-/i } keys %{$prog->db()};
	foreach my $opl ( @regopl ) {
	  # determine separate chunks
	  my @chunks;

	  if ( $details->{$opl} ) {
	    my $keuzepakket = $details->{KP} || 0;

	    if ( $keuzepakket ) {
	      my %keuzepks = map { m/(\d+)\s*-\s*(.+)/; $1 => $2 }
		split( /\s*,\s*/, $prog->db()->{$opl}->{KP} );

	      die( "Error: Keuzepakket '$keuzepakket' is unkown for OO $oo\n" )
		unless ( exists $keuzepks{$keuzepakket} );

	      @chunks = ( $keuzepks{$keuzepakket} );
	    } else {
	      my @minors = split( /\s*,\s*/, $prog->db()->{$opl}->{MINORS} );

	      # determine if the $oo is a major or a minor $oo
	      my $type;
	      my @minorkeys = grep { m/MINOR-/ } keys %$details;
	      if ( @minors == @minorkeys ) {
		@minorkeys = ( '' );
	      }
	      @chunks = @minorkeys;
	    }

	    foreach my $chunk ( @chunks ) {
	      $currdb
		->{$opl}
		  ->{$details->{PROGJAAR}}
		    ->{$chunk}
		      ->{$details->{SEM}}
			->{$details->{CAT}}
			  ->{$sgnr}
			    = $data;
	    }


	  }
	}

	my @shortopl = grep { m/^SP|^VP|^VT/i } keys %{$prog->db()};
	foreach my $opl ( @shortopl ) {
	  if ( $details->{$opl} ) {
	    use integer;
	    my $semester;
	    my $progjaar;
	    if ( $details->{$opl} < 10 ) {
	      $progjaar = ($details->{$opl} + 1) / 2;
	      $semester = $details->{$opl};
	    } else {
	      $semester = $details->{$opl};
	      $progjaar = 1 if( $semester == 12 );
	      $progjaar = 2 if( $semester == 34 );
	    }
	    $currdb
	      ->{$opl}
		->{$progjaar}
		  ->{''}
		    ->{$semester} # for semester we take the marked value!
		      ->{$details->{CAT}}
			->{$sgnr}
			  = $data;
	  }
	}
      }
    }
  }

  ConLogger::logitem( 'Generating curriculum tables' );

  foreach my $opl ( sort keys %{$currdb} ) {
    # open new file

    my $oFileName = File::Spec->catfile( $path, "Curriculum-IW-$opl.xlsx" );
    my $currbook =
      XLDB::OFile->new( filename => $oFileName, 
			title    => 'FacAdmin - Opdrachten',
			author   => 'Walter Daems',
			manager  => 'Walter Daems',
			company  => 'Universiteit Antwerpen',
			division => 'Faculteit Toegepaste Ingenieurswetenschappen',
			toolname => 'DMW - FacAdmin' );

    my $oplblurb = $prog->db()->{$opl}->{FULLNAME};

    foreach my $jaar ( sort keys %{$currdb->{$opl}} ) {
      # add sheet
      my $jaarblurb = $prog->db()->{$opl}->{TYPE} . "-" . $jaar;

      my $ysh = $currbook->makeSheet( $jaarblurb, 'L' );

      # add header
      my $row = 0;
      my $col = 0;
      $ysh->write( $row++, $col, "Opleiding: $oplblurb", 'b' );
      $ysh->write( $row++, $col, "Modeltraject: $jaarblurb" );

      $ysh->set_column( $col, $col, 16, 'L' );
      $ysh->write( $row, $col++, "Studiegidsnummer", 'Lx' );

      $ysh->set_column( $col, $col, 16, 'L' );
      $ysh->write( $row, $col++, "Module", 'Lx' );

      $ysh->set_column( $col, $col, 48, 'L' );
      $ysh->write( $row, $col++, "Opleidingsonderdeel", 'Lx' );

      $ysh->set_column( $col, $col, 5, 'RI1' );
      $ysh->write( $row, $col++, "SP", 'RI1x' );
      $ysh->set_column( $col, $col, 5, 'RI1' );
      $ysh->write( $row, $col++, "T", 'RI1x' );
      $ysh->set_column( $col, $col, 5, 'RI1' );
      $ysh->write( $row, $col++, "P", 'RI1x' );
      $ysh->set_column( $col, $col, 5, 'RI1' );
      $ysh->write( $row, $col++, "I", 'RI1x' );

      $ysh->set_column( $col, $col, 60, 'L' );
      $ysh->write( $row, $col++, "Titularissen", 'Lx' );

      $ysh->set_column( $col, $col, 5, 'C' );
      $ysh->write( $row, $col++, "Sem", 'Cx' );

      $ysh->set_column( $col, $col, 5, 'C' );
      $ysh->write( $row, $col++, "D", 'Cx' );

      $ysh->set_column( $col, $col, 5, 'C' );
      $ysh->write( $row, $col++, "ExCo", 'Cx' );
      ++$row;

      foreach my $minor ( sort keys %{$currdb->{$opl}->{$jaar}} ) {
	if ( length( $minor ) ) {
	  my $col = 0;
	  $ysh->write ( $row++, $col, $minor, 'B' );
	}
	foreach my $sem ( sort keys %{$currdb->{$opl}->{$jaar}->{$minor}} ) {
	  # write block
	  foreach my $cat ( sort keys %{$currdb->{$opl}->{$jaar}->{$minor}->{$sem}} ) {
	    foreach my $sgnr ( sort keys %{$currdb->{$opl}->{$jaar}->{$minor}->{$sem}->{$cat}} ) {
	      my $data = $currdb->{$opl}->{$jaar}->{$minor}->{$sem}->{$cat}->{$sgnr};
	      # write line
	      $col = 0;
	      $ysh->write( $row, $col++, $sgnr );
	      $ysh->write( $row, $col++, $cat );
	      $ysh->write( $row, $col++, $data->{OO} );
	      $ysh->write( $row, $col++, $data->{SP} );
	      $ysh->write( $row, $col, $data->{HC} ) if( $data->{HC} );
	      $ysh->write( $row, $col+1, $data->{PR} ) if( $data->{PR} );
	      $ysh->write( $row, $col+2, $data->{I} ) if( $data->{I} );
	      $col += 3;
	      $ysh->write( $row, $col++, $data->{TIT}, 'LW' );
	      $ysh->write( $row, $col++, $sem );
	      $ysh->write( $row, $col++, $data->{DEELTIJDS} );
	      ++$row
	    }
	  }
	  # skip line
	}
	++$row;
      }
      --$row;
      for ( $col = 0; $col <= 10; ++$col ) {
	$ysh->write($row, $col, undef, 'Cx' );
      }
      $ysh->set_row( $row, 4 );
      $row += 2;

      $col = 0;

      my $comment = $prog->db()->{$opl}->{COMMENT};
      $ysh->write( $row, $col++, "Opmerking:", 'B' );
      $ysh->write( $row++, $col, $comment );
      $row++;

      $col = 0;
      $ysh->write( $row++, $col, "Module: verzameling van opleidingsonderdelen in te vullen door de faculteit" );
      $ysh->write( $row++, $col, "T = aantal contacturen theorie, P = aantal contacturen praktijk, I = andere werkvormen (oefeningen, coaching, PGO, ..." );
      $ysh->write( $row++, $col, "SP = aantal studiepunten" );
      $ysh->write( $row++, $col, "Sem = semester" );
      $ysh->write( $row++, $col, "D = deeltijds programma (1 = deel 1, 2 = deel 2) " );
      $ysh->write( $row++, $col, "ExCo = 'x' indien niet te volgen onder examencontract" );
    }
    # close file
    $currbook->close();
  }
}

if ( $roster ) {
  ConLogger::logitem('Generating roster database' );

  my $rosterdb = {};
  while ( my ( $sgnr, $cro ) = each( %{$einv->db()} ) ) {

    my $c = $cro->{C};
    my $headmaster; # hoofdtitularis van OO
  CC:
    while ( my ( $oo, $fields ) = each( %$c ) ) {
      # loop will only execute once because of check of duplicate SGNRs
      $headmaster = $fields->{DOCENT};
      $rosterdb->{$headmaster}->{$oo}->{$sgnr}->{C} = $fields;
      $rosterdb->{ALL}->{$oo}->{$sgnr}->{C} = $fields;
    }

    my $r = $cro->{R};
  RR:
    while ( my ( $oo, $activities ) = each( %$r ) ) {
      # loop will only execute once because of check of duplicate SGNRs
      while ( my ( $act, $details ) = each( %$activities ) ) {
	my $noopcount = 0;
	while( my ( $group, $fields ) = each ( %$details ) ) {
	  $rosterdb->{$headmaster}->{$oo}->{$sgnr}->{R}->{$act}->{$group} = $fields;
	  $rosterdb->{ALL}->{$oo}->{$sgnr}->{R}->{$act}->{$group} = $fields;
	}
      }
    }
  }

  ########################
  # Write sample mailbody

  ConLogger::logitem( 'Writing sample mailbody' );

  my $mailbodyfilename = File::Spec->catfile( $path, "Roster-mailbody.txt" );
  my $mailbody = IO::File->new();
  $mailbody->open( ">$mailbodyfilename" )
    or die( "Error: cannot open mail body file '$mailbodyfilename' for writing\n" );
  print $mailbody <<EOF;
Geachte <MM:Callname>
Beste <MM:Firstname>

Je bent hoofdtitularis van een aantal opleidingsonderdelen.
Als bijlage vind je een excelbestand (<MM:ATTACHMENTS>) met een overzicht van deze opleidingsonderdelen en de gegevens waarop het uurrooster zal gebaseerd worden.

Ik wil je vragen deze gegevens grondig na te kijken en te corrigeren indien nodig.
Corrigeren kan met potlood/pen op een afdruk op papier, waarna je het blad aan mij of aan Lut Gulickx terugbezorgt.
Corrigeren kan ook in het excelbestand zelf. Geef de cellen die je wijzigt dan een gele achtergrondkleur. Stuur het bestand dan per e-mail naar walter.daems\@ua.ac.be.

Denk je dat je onterecht als hoofdtitularis voor een opleidingsonderdeel genoteerd staat, gelieve dit dan onmiddellijk te melden.

Velden waar 'XXX' of een vraagteken instaan duiden ontbrekende gegevens aan.

Wat lokalen betreft:
  - lokalen van de campus Paardenmarkt starten met de letters 'PM.'
  - lokalen van de campus Hoboken starten met de letters 'H.'
  - lokalen van de campus Groenenborger starten met de letters 'G.'
  - lokalen van de stadscampus starten met de letters 'S.'
Dit is conform het officiële lokalennaamgevingsschema van de Universiteit Antwerpen.
'A' duidt een aula aan, 'O' een oefeningenlokaal. Een PC-lokaal kan je aanduiden met PC-labo of PC-klas.

De omschrijving tussen haakjes gebruikt afkortingen BA-YYY, MA-YYY, SP-YYY, VP-YYY, VT-YYY waarbij YYY een inhoudelijk logische omschrijving is. BA staat voor bachelor, MA voor master, SP voor schakelprogramma, VP voor voorbereidingsprogramma en VT voor verkort traject.

De groepsnummering vermeld in de kolom 'Detail' is nog niet finaal. Dat is op zich geen probleem. Wat wel correct _moet_ zijn is het aantal groepen in combinatie met aantal sessies en sessieduur. Merk op dat voor het nieuwe curriculum én op de campus Paardenmarkt enkel een sessieduur van 2u of veelvouden van 2u is toegelaten.

Je correcties (of een boodschap dat alles in orde is) moet ons tijdig bereiken om het rooster iniet in gevaar te brengen.
Deadline: dinsdag 26 maart 2013 om 23:59

Ik ben de volgende dagen beschikbaar om jullie bij de controle te helpen en indien mogelijk ineens jullie wijzigingen te verwerken:
  - woensdag 20/3 van 11u-18u op de campus Hoboken (docentenruimte)
  - vrijdag 22/3 van 11u-18u op de campus Paardenmarkt (docentenruimte)
  - maandag 24/3 van 9u-12u30 op de campus Paardenmarkt (docentenruimte)
  - dinsdag 26/3 van 9u-12u30 op de campus Hoboken (docentenruimte)


Met vriendelijke groeten


Walter Daems
Academisch Faculteitscoördinator
EOF

  ############################
  # Write mailmerge excelbase

  ConLogger::logitem( 'Writing mailmerge excelbase' );

  my $mmFileName = File::Spec->catfile( $path, "Roster-mailmerge.xlsx" );
  my $mmbook = 
    XLDB::OFile->new( filename => $mmFileName,
		      title    => 'FacAdmin - Roster - Mailmerge',
		      author   => 'Walter Daems',
		      manager  => 'Walter Daems',
		      company  => 'Universiteit Antwerpen',
		      division => 'Faculteit Toegepaste Ingenieurswetenschappen',
		      toolname => 'DMW - FacAdmin' );


  my $mmsh = $mmbook->makeSheet( "Mailmerge data", 'P' );

  my $mmrow = 0;
  my $mmcol = 0;
  $mmsh->write( $mmrow, $mmcol, "Mailmerge data", 'B' );

  ++$mmrow;
  $mmsh->write( ++$mmrow, $mmcol, "Global values" , 'Lx' );
  $mmsh->write( $mmrow, $mmcol+1, "(ADD / DEFAULT / IGNORE / OVERRULE)" , 'Lx' );
  $mmsh->write( $mmrow, $mmcol+2, undef , 'Lx' );

  $mmsh->write( ++$mmrow, $mmcol, "FROM", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "OVERRULE" );
  $mmsh->write( $mmrow, $mmcol+2, "walter.daems\@ua.ac.be" );

  $mmsh->write( ++$mmrow, $mmcol, "TO", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "OVERRULE" );
  $mmsh->write( $mmrow, $mmcol+2, "walter.daems\@ua.ac.be" );

  $mmsh->write( ++$mmrow, $mmcol, "CC", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "IGNORE" );
  $mmsh->write( $mmrow, $mmcol+2, undef );

  $mmsh->write( ++$mmrow, $mmcol, "BCC", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "IGNORE" );
  $mmsh->write( $mmrow, $mmcol+2, undef );

  $mmsh->write( ++$mmrow, $mmcol, "SUBJECT", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "OVERRULE" );
  $mmsh->write( $mmrow, $mmcol+2, "DRINGEND: Controle roosterinformatie" );

  $mmsh->write( ++$mmrow, $mmcol, "BODY", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "OVERRULE" );
  $mmsh->write( $mmrow, $mmcol+2, $mailbodyfilename, 'L' );

  $mmsh->write( ++$mmrow, $mmcol, "ATTACHMENTS", 'B' );
  $mmsh->write( $mmrow, $mmcol+1, "IGNORE" );

  $mmsh->write( $mmrow, $mmcol+2, undef );

  ++$mmrow;
  $mmsh->write( ++$mmrow, $mmcol, "Detailed values", 'Lx' );
  $mmsh->write( $mmrow, $mmcol+1, undef , 'Lx' );
  $mmsh->write( $mmrow, $mmcol+2, undef , 'Lx' );

  $mmsh->write( ++$mmrow, $mmcol, "TO", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "CC", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "BCC", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "SUBJECT", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "BODY", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "ATTACHMENTS", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "Lastname", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "Firstname", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "Fullname", 'B' );
  $mmsh->write(   $mmrow, ++$mmcol, "Callname", 'B' );
  $mmsh->set_column( 0, $mmcol, 25, 'L' );

  #################################
  # Write data for each headmaster

  ConLogger::logitem( 'Writing roster files' );

  {
    my $pb = ConLogger::ProgressBar->new();
    my $hdcount = keys %{$rosterdb};
    my $hdcounter = 0;

    foreach my $headmaster ( sort keys %{$rosterdb} ) {
      $pb->progress( $hdcounter++, $hdcount );

      my $oFileName = File::Spec->catfile( $path, "Roster-$headmaster.xlsx" );

      # complete mailmerge form
      my $persdata = $pers->db()->{$headmaster};

      $mmcol = 0;
      $mmsh->write( ++$mmrow,   $mmcol, $persdata->{EMAIL} );
      $mmcol += 4;
      $mmsh->write( $mmrow, ++$mmcol, $oFileName );
      $mmsh->write( $mmrow, ++$mmcol, $persdata->{LASTNAME} );
      $mmsh->write( $mmrow, ++$mmcol, $persdata->{FIRSTNAME} );
      $mmsh->write( $mmrow, ++$mmcol, $headmaster );
      $mmsh->write( $mmrow, ++$mmcol, "Collega" );

      # open new file
      my $currbook =
	XLDB::OFile->new( filename => $oFileName,
			  title    => 'FacAdmin - Roster - $headmaster',
			  author   => 'Walter Daems',
			  manager  => 'Walter Daems',
			  company  => 'Universiteit Antwerpen',
			  division => 'Faculteit Toegepaste Ingenieurswetenschappen',
			  toolname => 'DMW - FacAdmin' );

      my $sh = $currbook->makeSheet( "Roster", 'L' );

      my $row = 0;
      my $col = 0;
      $sh->write( $row++, $col, "Roosteroverzicht ter controle voor " . $headmaster . " (hoofdtitularis)",
		  'B' );

      $sh->write( ++$row, $col, "Stugiegidsnr", 'Lx' );
      $sh->set_column( $col, $col, 12, 'L' );
      $sh->write( $row, ++$col, "Prog", 'Lx' );
      $sh->set_column( $col, $col, 5, 'RI1' );
      $sh->write( $row, ++$col, "Yr", 'Lx' );
      $sh->set_column( $col, $col, 3, 'RI1' );
      $sh->write( $row, ++$col, "Opleidingsonderdeel > Activiteit", 'Lx' );
      $sh->set_column( $col, $col, 40, 'L' );
      $sh->write( $row, ++$col, "CU", 'Rx' );
      $sh->set_column( $col, $col, 4, 'R' );
      $sh->write( $row, ++$col, "TP", 'Rx' );
      $sh->set_column( $col, $col, 4, 'R' );
      $sh->write( $row, ++$col, "Sem", 'Lx' );
      $sh->set_column( $col, $col, 5, 'RI1' );
      $sh->write( $row, ++$col, "Docenten", 'Lx' );
      $sh->set_column( $col, $col, 25, 'LW' );
      $sh->write( $row, ++$col, "#Groepen", 'Lx' );
      $sh->set_column( $col, $col, 10, 'RI1' );
      $sh->write( $row, ++$col, "Detail", 'Lx' );
      $sh->set_column( $col, $col, 20, 'L' );
      $sh->write( $row, ++$col, "N sessies", 'Cx' );
      $sh->set_column( $col, $col, 8, 'RI1' );
      $sh->write( $row, ++$col, "van Xh", 'Cx' );
      $sh->set_column( $col, $col, 8, 'RI1' );
      $sh->write( $row, ++$col, "Campus", 'Lx' );
      $sh->set_column( $col, $col, 10, 'L' );
      $sh->write( $row, ++$col, "Lokaal", 'Lx' );
      $sh->set_column( $col, $col, 10, 'L' );
      $sh->write( $row, ++$col, "Opmerking", 'Lx' );
      $sh->set_column( $col, $col, 50, 'LW' );
      my $endcol = $col;

      my $oodb = $rosterdb->{$headmaster};
      foreach my $oo ( sort keys %{$oodb} ) {
	my $sgnrdb = $oodb->{$oo};
	foreach my $sgnr ( sort keys %{$sgnrdb} ) {
	  my $col = 0;
	  $sh->write( ++$row, $col, $sgnr, 'Lb' );
	  $sh->write( $row, ++$col, $sgnrdb->{$sgnr}->{C}->{PROG}, 'RIb' );
	  $sh->write( $row, ++$col, $sgnrdb->{$sgnr}->{C}->{PROGJAAR}, 'RIb' );
	  $sh->write( $row, ++$col, $oo . " (" . $sgnrdb->{$sgnr}->{C}->{SP} . " ECTS)", 'Lb' );
	  my $oocol = $col;
	  $sh->write( $row, $col+=3, $sgnrdb->{$sgnr}->{C}->{SEM}, 'RIb' );
	  $sh->write( $row, ++$col, $sgnrdb->{$sgnr}->{C}->{DOCENT}, 'Lb' );

	  my $curriculumstring = "(OO in curriculum van ";
	  foreach my $currkey (sort grep { m/^BA-|^MA-|^VP-|^VT|^SP-/ } keys $sgnrdb->{$sgnr}->{C} ) {
	    $curriculumstring .= $currkey . "/" if ( $sgnrdb->{$sgnr}->{C}->{$currkey} );
	  }
	  chop $curriculumstring;
	  $curriculumstring .= ")";
	  $sh->write( $row, ++$col, $curriculumstring, 'Lb' );

	  my $actdb = $sgnrdb->{$sgnr}->{R};
	  ++$row;
	  foreach my $act ( sort keys %{$actdb} ) {
	    $col = $oocol;
	    my $grpdb = $actdb->{$act};
	    $sh->write( $row, $col, "  > " . $act );
	    my $grpcol = ++$col;
	    my $grpcnt = 0;
	    foreach my $grp ( sort keys %{$grpdb} ) {
	      $col = $grpcol;
	      my $grpact = $grpdb->{$grp};
	      if ( $grpcnt ) {
		++$col;
	      } else {
		$sh->write( $row, $col,   $grpact->{CUTYPE} );
		$sh->write( $row, ++$col, $grpact->{CU} . "h" );
	      }
	      ++$grpcnt;
	      $sh->write( $row, ++$col, $grpact->{SEM} );
	      $sh->write( $row, ++$col, $grpact->{DOCENT} );
	      $sh->write( $row, ++$col, $grpact->{AGR} );
	      $sh->write( $row, ++$col, $grp );
	      $sh->write( $row, ++$col, $grpact->{ASS} );
	      $sh->write( $row, ++$col, $grpact->{DSS} . "h" );
	      $sh->write( $row, ++$col, $grpact->{CAMPUS} );
	      $sh->write( $row, ++$col, $grpact->{LOKAAL} );
	      $sh->write( $row, ++$col, $grpact->{OPM} );
	      ++$row;
	    }
	  }
	}
      }
      for ( my $i = 0; $i <= $endcol; ++$i ) {
	$sh->write( $row, $i, undef, 'Lx' );
      }
      $sh->set_row( $row, 4 );

      $currbook->close();
    }
  }
  $mmbook->close();
}


sub findActiveCUTag {
  my ( $detail ) = @_;
  foreach my $tag ( keys %$detail ) {
    next if( $tag =~ /^AGR|CAMPUS|GGR|SEM|AGR|NEW|PROG|PROGJAAR$/ );
    return $tag;
  }
}


sub labelof {
  my ( undef, $label )= split( /\|/, $_[0] );
  return $label;
}

__END__


=head1 NAME

facAdmin - perform faculty administration

=head1 SYNOPSIS

facAdmin [options] [files ...]

 Options:
   --help            brief help message
   --man             full documentation

=head1 OPTIONS

=over 8

=item B<--help>

Print a brief help message and exits.

=item B<--man>

Prints the manual page and exits.

=back

=head1 DESCRIPTION

B<This program> will read the given input file(s) and do something
useful with the contents thereof.

=cut
