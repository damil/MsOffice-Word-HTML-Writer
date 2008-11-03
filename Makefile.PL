use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'MsOffice::Word::HTML::Writer',
    AUTHOR              => 'Laurent Dami <laurent.dami AT etat.ge.ch>',
    VERSION_FROM        => 'lib/MsOffice/Word/HTML/Writer.pm',
    ABSTRACT_FROM       => 'lib/MsOffice/Word/HTML/Writer.pm',
    PL_FILES            => {},
    PREREQ_PM => {
        'Test::More' => 0,
        'MIME::QuotedPrint' => 0,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'MsOffice-Word-HTML-Writer-*' },
);
