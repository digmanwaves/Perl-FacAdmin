# -*- cperl -*-
# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

my $cntr = 1;
BEGIN { $| = 1; print "1..3\n"; }
END { print "not ok $cntr\n" unless $cntr == 0 }
use XLDB::IFile;
use FacAdmin::EduInventory;
print "ok $cntr\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

# check method new
$cntr = 2;
my $xldbfile = XLDB::IFile->new();
$xldbfile->open( 't/EduInventory.xlsm' );
my $sheet = $xldbfile->sheet( 'Onderwijsoverzicht' );
my $progdb = FacAdmin::EduInventory->new();
$progdb->connect( $sheet );
print "ok $cntr\n";

# check method buildDataBase
$cntr = 3;
$progdb->buildDataBase();

# my $expectedprogs = "SP-PBAUe-MAEI|SP-PBBL-MABC|SP-PBCHb-MABC|SP-PBCHcp-MACH|SP-PBEI-MAEI|SP-PBEMAU-MAEM|SP-PBMCT-MAEI|SP-PBTI-MAEI|VP-MAAR-MABK|VP-MAHI-MAEM|VP-MANW-MAEM|VT-ABAR-MABK|VT-PBBK-MABK|VT-PBVG-MABK";
# my $detectedprogs = join( "|", sort keys %{$progdb->db()} );
# die( "Expected: $expectedprogs; Got: $detectedprogs\n" )
#   unless ( $expectedprogs eq $detectedprogs );

# foreach my $property ( qw ( MINORS FULLNAME COMMENT ACRONYM TYPE KP AR ) ) {
#   die( "Did not find property '$property' on 'VT-ABAR-MABK'\n" )
#     unless( exists $progdb->db()->{'VT-ABAR-MABK'}->{$property} );
# }
print "ok $cntr\n";

$cntr = 0;
