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
use FacAdmin::RptSt;
use Data::Dumper;
print "ok $cntr\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

# check method new
$cntr = 2;
my $xldbfile = XLDB::IFile->new();
$xldbfile->open( 't/QQ_RPTST_1429016.xlsx' );
my $sheet = $xldbfile->sheet( 'Sheet1' );
my $rptst = FacAdmin::RptSt->new();
$rptst->connect( $sheet );
print "ok $cntr\n";


# check method buildDataBase
$cntr = 3;
$rptst->buildDataBase();
#print STDERR Data::Dumper->Dump( [ $rptst ], [ "Report Study Guide" ] );
print "ok $cntr\n";


$cntr = 0;
