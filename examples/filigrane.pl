use strict;
use warnings;
use lib "../lib";

use MsOffice::Word::HTML::Writer;

my $doc = MsOffice::Word::HTML::Writer->new(
  title => "My new doc",
 );

$doc->break_section(
  first_header => sprintf("<P>THIS IS THE FIRST HEADER</P>, page %s of %s at %s</P>%s",
                           $doc->field('PAGE'),
                           $doc->field('NUMPAGES') ,
                           $doc->field('DATE', q{\\@ "dddd d MMMM yyyy" \\* Upper}),
                           filigrane()),
  header       => sprintf ("<P>THIS IS THE HEADER, page %s of %s</P>",
                           $doc->field('PAGE'),
                           $doc->field('NUMPAGES')) ,


  first_footer => "<P>THIS IS THE FIRST FOOTER</P>",
  footer       => "<P>THIS IS THE FOOTER</P>",
);


$doc->write(<<"");
   <P>HELLO, WORLD</P>
   <P>THIS IS A GENERATED DOC</P>

$doc->write(filigrane());


$doc->break_page;

$doc->write(<<"");
   <p>Page two here</p>


$doc->save_as("foo");

system("start foo.doc");


sub filigrane {
  return <<"";
<span>
<v:shape style='position:absolute;left:5cm;top:5cm;width:429.75pt;height:96pt;rotation:-3191716fd;'>
 <v:textpath style='font-family:"Arial Black";v-text-kern:t' trim="t"
  fitpath="t" string="DUPLICATA" />
</v:shape>
</span>

}
