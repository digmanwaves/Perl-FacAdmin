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
use FacAdmin::Personnel;
print "ok $cntr\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

# check method new
$cntr = 2;
my $xldbfile = XLDB::IFile->new();
$xldbfile->open( 't/Personnel.xlsx' );
my $sheet = $xldbfile->sheet( 'Personeelsoverzicht' );
my $persdb = FacAdmin::Personnel->new();
$persdb->connect( $sheet );
print "ok $cntr\n";

# check method buildDataBase
$cntr = 3;
$persdb->buildDataBase();
my $expectedpersons = "Bhoy Danny|Bindozer Walter|Dumbhead Dirk|Frog Frank|GI Joe|Ketten Dick|MacAllen Bob|Sporgese Ellen";
my $detectedpersons = join( "|", sort keys %{$persdb->db()} );
die( "Expected: $expectedpersons; Got: $detectedpersons\n" )
  unless ( $expectedpersons eq $detectedpersons );

foreach my $property ( qw ( LASTNAME FIRSTNAME OPL VTE OZP DVP ) ) {
  die( "Did not find property '$property' on 'Danny Bhoy'\n" )
    unless( exists $persdb->db()->{'Bhoy Danny'}->{$property} );
}
print "ok $cntr\n";


$cntr = 0;
