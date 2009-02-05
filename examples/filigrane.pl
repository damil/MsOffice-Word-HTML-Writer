use strict;
use warnings;
use lib "../lib";

use MsOffice::Word::HTML::Writer;

my $doc = MsOffice::Word::HTML::Writer->new(
  title => "My new doc",
 );

$doc->create_section(
  page => {size          => "21.0cm 29.7cm",
           margin        => "1.2cm 2.3cm 4.5cm 6.7cm",
           header_margin => "4cm"},
  first_header => sprintf("<P>THIS IS THE FIRST HEADER</P>, page %s of %s at %s</P>%s"
                             . filigrane("HELLO FIRST"),

                           $doc->field('PAGE'),
                           $doc->field('NUMPAGES') ,
                           $doc->field('DATE', q{\\@ "dddd d MMMM yyyy" \\* Upper})
                           ),
  header       => sprintf ("<P>THIS IS THE HEADER, page %s of %s</P>",
                           $doc->field('PAGE'),
                           $doc->field('NUMPAGES')) ,


  first_footer => "<P>THIS IS THE FIRST FOOTER</P>",
  footer       => "<P>THIS IS THE FOOTER</P>",
);


$doc->write(<<"");
   <p style='margin-left:-2.0cm'>HELLO, WORLD</P>
   <P>THIS IS A GENERATED DOC</P>
   <p>prions d&#8217;agréer, Monsieur</p>


# $doc->write(filigrane());


$doc->write($doc->page_break);

$doc->write(<<"");
   <p>Page two here</p>





$doc->save_as("foobar");

system("start foobar.doc");


sub filigrane {
  my $text = shift;

  my $template = q{
<v:shapetype id="filigrane_type"
   coordsize="21600,21600" 
   adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
 <v:formulas>
  <v:f eqn="sum #0 0 10800"/>
  <v:f eqn="prod #0 2 1"/>
  <v:f eqn="sum 21600 0 @1"/>
  <v:f eqn="sum 0 0 @2"/>
  <v:f eqn="sum 21600 0 @3"/>
  <v:f eqn="if @0 @3 0"/>
  <v:f eqn="if @0 21600 @1"/>
  <v:f eqn="if @0 0 @2"/>
  <v:f eqn="if @0 @4 21600"/>
  <v:f eqn="mid @5 @6"/>
  <v:f eqn="mid @8 @5"/>
  <v:f eqn="mid @7 @8"/>
  <v:f eqn="mid @6 @7"/>
  <v:f eqn="sum @6 0 @5"/>
 </v:formulas>
 <v:path textpathok="t" 
    o:connecttype="custom" 
    o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"
    o:connectangles="270,180,90,0"/>
 <v:textpath on="t" fitshape="t"/>
 <v:handles>
  <v:h position="#0,bottomRight" xrange="6629,14971"/>
 </v:handles>
 <o:lock v:ext="edit" text="t" shapetype="t"/>
</v:shapetype>

<v:shape id="foobar_filigrane" type="#filigrane_type"
 style='position:absolute;margin-left:0;margin-top:0;
    width:429.75pt;height:96pt;
    rotation:315;z-index:2;mso-position-horizontal:center;
    mso-position-horizontal-relative:margin;mso-position-vertical:center;
    mso-position-vertical-relative:margin' 
 fillcolor="silver"
 stroked="f">
 <v:fill opacity=".5"/>
 <v:textpath style='font-family:"Arial Black"' string="%s"/>
 <w:wrap anchorx="margin" anchory="margin"/>
</v:shape>
};

  return sprintf $template, $text;
}
