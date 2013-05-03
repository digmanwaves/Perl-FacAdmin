require 5.006;

use ExtUtils::MakeMaker;

my $target = 'MS'; # set this to MS to avoid generating man pages with 

@scripts = qw( facAdmin.pl );

WriteMakefile
  (
   NAME      => 'FacAdmin',
   AUTHOR    => 'Walter Daems / Digital Manifold Waves (walter@digmanwaves.net)',
   ABSTRACT  => 'Faculty Administration Suite based on XLDB',
   VERSION   => '1.00',
   PREREQ_PM => { 'XLDB::File' => 0,
                  'XLDB::Sheet' => 0,
		  'ConLogger' => 0,
		  'ConLogger::SubTask' => 0,
   },
   LICENSE   => 'perl',
   EXE_FILES => [ map { "bin/$_" } @scripts ],
   ( $target eq 'MS' ) ? ( MAN1PODS => {}, MAN3PODS => {} ) : (),
   test      => { TESTS => 't/???-test-*.pl' },
);


