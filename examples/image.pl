use strict;
use warnings;
use lib "../lib";

use MsOffice::Word::HTML::Writer;

my $doc = MsOffice::Word::HTML::Writer->new(
  title => "Test embedded image",
 );


use File::Slurp;
my $img = read_file("ecuge.png", binmode => ":raw");
$doc->add_MIME_part("ecuge.png" => $img);


$doc->write(<<"");
<img width=35 height=57 src="files/ecuge.png"></img>
   <P>HELLO, WORLD</P>
   <P>THIS IS A GENERATED DOC</P>

$doc->save_as("tst_image");

