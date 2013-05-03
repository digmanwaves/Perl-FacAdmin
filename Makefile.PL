require 5.006;

use ExtUtils::MakeMaker;

@scripts = qw( facAdmin.pl );

WriteMakefile
  (
   NAME      => 'FacAdmin',
   AUTHOR    => 'Walter Daems (walter.daems@ua.ac.be)',
   ABSTRACT  => 'Faculty Administration Suite based on XLDB',
   VERSION   => '1.00',
   PREREQ_PM => { 'XLDB::File' => 0,
                  'XLDB::Sheet' => 0,
		  'ConLogger' => 0,
		  'ConLogger::SubTask' => 0,
   },
   LICENSE   => 'perl',
   EXE_FILES => [ map { "bin/$_" } @scripts ],
   test      => { TESTS => 't/???-test-*.pl' },
);


