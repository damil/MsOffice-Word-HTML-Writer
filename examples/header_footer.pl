use strict;
use warnings;
use lib "../lib";

use MsOffice::Word::HTML::Writer;

my $doc = MsOffice::Word::HTML::Writer->new(
  title => "My new doc",
 );

$doc->create_section(
  first_header => sprintf("<P>THIS IS THE FIRST HEADER</P>, page %s of %s at %s</P>",
                           $doc->field('PAGE'),
                           $doc->field('NUMPAGES') ,
                           $doc->field('DATE', q{\\@ "dddd d MMMM yyyy" \\* Upper})),
  header       => sprintf ("<P>THIS IS THE HEADER, page %s of %s</P>",
                           $doc->field('PAGE'),
                           $doc->field('NUMPAGES')) ,


  first_footer => "<P>THIS IS THE FIRST FOOTER</P>",
  footer       => "<P>THIS IS THE FOOTER</P>",
);


$doc->write(<<"");
   <P>HELLO, WORLD</P>
   <P>THIS IS A GENERATED DOC</P>


$doc->page_break;

$doc->write(<<"");
   <p>Page two here</p>


$doc->save_as("foo");

system "start foo.doc";
