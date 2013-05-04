require 5.006;

use ExtUtils::MakeMaker;

$version = '1.00';

my $target = 'MS'; # set this to MS to avoid generating man pages with 

@scripts = qw( facAdmin.pl );

WriteMakefile
  (
   NAME      => 'FacAdmin',
   AUTHOR    => 'Walter Daems / Digital Manifold Waves (walter@digmanwaves.net)',
   ABSTRACT  => 'Faculty Administration Suite based on XLDB',
   VERSION   => $version,
   PREREQ_PM => { 'XLDB::File' => 0,
                  'XLDB::IFile' => 0,
                  'XLDB::OFile' => 0,
                  'XLDB::DataBase' => 0,
                  'XLDB::Sheet' => 0,
                  'ConLogger' => 0,
                  'ConLogger::SubTask' => 0,
                  'ConLogger::ProgressBar' => 0,
   },
   LICENSE   => 'perl',
   EXE_FILES => [ map { "bin/$_" } @scripts ],
   ( $target eq 'MS' ) ? ( MAN1PODS => {}, MAN3PODS => {} ) : (),
   META_ADD  => 
   { provides => {
       'FacAdmin::EduInventory' => { file    => 'lib/FacAdmin/EduInventory.pm',
				     version => $version },
       'FacAdmin::Personnel'    => { file    => 'lib/FacAdmin/Personnel.pm',
				     version => $version },
       'FacAdmin::Programmes'   => { file    => 'lib/FacAdmin/Programmes.pm',
				     version => $version },
       'FacAdmin::RptSt'        => { file    => 'lib/FacAdmin/RptSt.pm',
				     version => $version },
     }
   },
   test      => { TESTS => 't/???-test-*.pl' },
  );

