package MsOffice::Word::HTML::Writer;

use warnings;
use strict;
use MIME::QuotedPrint qw/encode_qp/;
use MIME::Base64      qw/encode_base64/;
use MIME::Types;
use Carp;
our $VERSION = '0.01';

sub new {
  my ($class, %params) = @_;

  # check parameters
  my @bad = grep {$_ !~ /^(title|styles|hf_styles)$/}
                 keys %params;
  croak "new: bad parameter: " . join(", ", @bad) if @bad;

  # create instance
  my $self = {
    MIME_parts => [],
    sections   => [{}],
    title      => $params{title}
               || "Document generated by MsOffice::Word::HTML::Writer",
    styles     => $params{styles}    || "",
    hf_styles  => $params{hf_styles} || ""
   };

  bless $self, $class;
}


sub break_section {
  my ($self, %params) = @_;

  # check parameters
  my @bad = grep {$_ !~ /^(header|footer|first_header|first_footer|new_page)$/}
                 keys %params;
  croak "break_section : bad parameter: " . join(", ", @bad) if @bad;

  # if first automatic section if empty, delete it
  $self->{sections} = []
    if scalar(@{$self->{sections}}) == 1 && !$self->{sections}[0]{content};

  # add the new section
  push @{$self->{sections}}, \%params;

}


sub write {
  my ($self, $html) = @_;

  # add html to current section content
  $self->{sections}[-1]{content} .= $html;
}



sub break_page {
  my ($self, $html) = @_;

  $self->{sections}[-1]{content} 
    .= qq{<br clear='all' style='page-break-before:always'>\n};
}



sub save_as {
  my ($self, $filename) = @_;

  # default extension is ".doc"
  $filename .= ".doc" unless $filename =~ /[^\/\\]+\.\w+$/;

  # open the file
  open my $fh, ">", $filename
    or croak "could not open >$filename: $!";

  # write content and close
  print $fh $self->content;
  close $fh;
}


sub attach {
  my ($self, $name, $open1, $open2, @other) = @_;

  # open a handle to the attachment (need to dispatch according to number
  # of args, because perlfunc/open() has complex prototyping behaviour)
  my $fh;
  if (@other)    { open $fh, $open1, $open2, @other or die $!; }
  elsif ($open2) { open $fh, $open1, $open2         or die $!; }
  else           { open $fh, $open1                 or die $!; }

  # slurp the content
  binmode($fh);
  local $/;
  my $attachment = <$fh>;

  # add the attachment (filename and content)
  push @{$self->{MIME_parts}}, ["files/$name", $attachment];
}


sub field {
  my ($self, $fieldname, $args) = @_;

  my $field;

  # when args : long form of field encoding
  if ($args) {
      $field = qq{<span style='mso-element:field-begin'></span>}
             . "$fieldname $args"
             . qq{<span style='mso-element:field-separator'></span>}
             . qq{<span style='mso-element:field-end'></span>};
  }
  # otherwise : short form of field encoding
  else {
    $field = qq{<span style='mso-field-code:" $fieldname "'></span>};
  }

  return $field;
}



sub content {
  my ($self) = @_;

  # separator for parts in MIME document
  my $boundary = qw/__NEXT_PART__/;

  # MIME multipart header
  my $mime = qq{MIME-Version: 1.0\n}
           . qq{Content-Type: multipart/related; boundary="$boundary"\n\n}
           . qq{MIME document generated by MsOffice::Word::HTML::Writer\n\n};


  # generate each part (main document must be first)
  my @parts = $self->_MIME_parts;
  my $filelist = $self->_filelist(@parts);
  for my $pair ($self->_main, @parts, $filelist) {
    my ($filename, $content) = @$pair;
    my $mime_type = MIME::Types->new->mimeTypeOf($filename);
    my ($encoding, $encode_sub) 
      = $mime_type =~ /^text|xml$/ ? ('quoted-printable', \&encode_qp    )
                                   : ('base64',           \&encode_base64);

    $mime .= qq{--$boundary\n}
          .  qq{Content-Location: file:///C:/foo/$filename\n}
          .  qq{Content-Transfer-Encoding: $encoding\n}
          .  qq{Content-Type: $mime_type\n\n}
          .  $encode_sub->($content) 
          . "\n";
  }

  # close last MIME part
  $mime .= "--$boundary--\n";

  return $mime;
}


#======================================================================
# PRIVATE METHODS
#======================================================================

sub _main {
  my ($self) = @_;

  # body : concatenate content from all sections
  my $body = "";
  my $i = 1;
  foreach my $section (@{$self->{sections}}) {

    # section break
    if ($i > 1) {
      my $break = $section->{new_page} ? 'always' : 'auto';
      my $style = qq{page-break-before:$break;mso-break-type:section-break};
      $body .= qq{<br clear=all style='$style'>\n};
    }

    # section content
    $body .= qq{<div class="Section$i">\n$section->{content}\n</div>\n};

    $i += 1;
  }

  # assemble head and body into a full document
  my $html
    = qq{<html xmlns:v="urn:schemas-microsoft-com:vml"\n}
    . qq{      xmlns:o="urn:schemas-microsoft-com:office:office"\n}
    . qq{      xmlns:w="urn:schemas-microsoft-com:office:word"\n}
    . qq{      xmlns="http://www.w3.org/TR/REC-html40">\n}
    . $self->_head
    . qq{<body>\n$body</body>\n}
    . qq{</html>\n};
  return ["main.htm", $html];
}


sub _head {
  my ($self) = @_;

  # HTML head : link to filelist, title, view format and styles
  my $head 
    = qq{<head>\n}
    . qq{<link rel=File-List href="files/filelist.xml">\n}
    . qq{<title>$self->{title}</title>\n}
    . qq{<xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml>\n}
    . qq{<style>\n} . $self->_styles . qq{</style>\n}
    . qq{</head>\n};
  return $head;
}


sub _styles {
  my ($self) = @_;

  # styles supplied by user
  my $styles = $self->{styles} || "";

  # additional styles for linking to headers and footers in each section
  my $i = 1;
  foreach my $section (@{$self->{sections}}) {

    # which headers and footers have been defined in this section
    my $properties = "";
    my $has_first_page;
    foreach my $hf (qw/header footer first_header first_footer/) {
      $section->{$hf} or next;
      $has_first_page = 1 if $hf =~ /^first/;
      (my $property = $hf) =~ s/_/-/;
      $properties 
        .= qq{  mso-$property:url("files/header_footer.htm") $hf$i;\n};
    }
    $properties .= qq{mso-title-page:yes;\n} if $has_first_page;

    # style definitions for this section
    $styles .= qq[\@page Section$i {\n$properties}\n]
            .  qq[div.Section$i {page:Section$i}\n];
    $i += 1;
  }

  return $styles;
}


sub _MIME_parts {
  my ($self) = @_;

  # attachments supplied by user
  my @parts = @{$self->{MIME_parts}};

  # additional attachment : computed file with headers and footers
  my $hf_content = $self->_header_footer;
  unshift @parts, ["files/header_footer.htm", $hf_content] if $hf_content;

  return @parts;
}


sub _header_footer {
  my ($self) = @_;

  # create a div for each header/footer in each section
  my $hf_divs = "";
  my $i = 1;
  foreach my $section (@{$self->{sections}}) {

    # deal with headers/footers defined in that section
    foreach my $hf (qw/header footer first_header first_footer/) {
      $section->{$hf} or next;
      (my $style = $hf) =~ s/^first_//;
      $hf_divs .= qq{<div style='mso-element:$style' id='$hf$i'>\n}
               .  $section->{$hf} . "\n"
               .  qq{</div>\n};
    }

    $i += 1;
  }

  # if at least one such div, need to create an attached file
  $hf_divs 
    &&= qq{<html>\n}
      . qq{<head>\n}
      . qq{<link id=Main-File rel=Main-File href="../main.htm">\n}
      . qq{<style>\n} . $self->{hf_styles} . qq{</style>\n}
      . qq{</head>\n}
      . qq{<body>\n} . $hf_divs . qq{</body>\n}
      . qq{</html>\n};

  return $hf_divs;
}



sub _filelist {
  my ($self, @parts) = @_;

  # xml header
  my $xml = qq{<xml xmlns:o="urn:schemas-microsoft-com:office:office">\n}
          . qq{ <o:MainFile HRef="../main.htm"/>\n};

  # refer to each attached file
  foreach my $part (@parts) {
    $xml .= qq{ <o:File HRef="$part->[0]"/>\n};
  }

  # the filelist is itself an attached file
  $xml .= qq{ <o:File HRef="filelist.xml"/>\n};

  # closing tag;
  $xml .=  qq{</xml>\n};

  return ["files/filelist.xml", $xml];
}





1;

__END__

=head1 NAME

MsOffice::Word::HTML::Writer - Writing documents for MsWord in HTML format

=head1 SYNOPSIS

  use MsOffice::Word::HTML::Writer;
  my $doc = MsOffice::Word::HTML::Writer->new(title  => "My new doc");
  
  $doc->write("<p>hello, world</p>");
  
  $doc->break_section(
    header => sprintf("Section 2, page %s of %s", 
                                  $doc->field('PAGE'), 
                                  $doc->field('NUMPAGES')),
    footer => sprintf("printed at %s", 
                                  $doc->field('PRINTDATE')),
    new_page => 1,
  );
  $doc->write("this is the second section, look at header/footer");

  $doc->attach("my_image.gif", $path_to_my_image);
  $doc->write("<img src='files/my_image.gif'>");

  $doc->save_as("c:/tmp/my_doc");

=head1 DESCRIPTION

The present module is one way to programatically generate documents
targeted for Microsoft Word (MsWord).

MsWord can read documents encoded in native binary format, in Rich
Text Format (RTF), in WordML (an XML dialect), or -- maybe this is
less known -- in HTML, with some special markup for pagination and
other MsWord-specific features. Such HTML documents are often in
several parts, because attachments like images or headers/footers need
to be in separate files; however, since it is more convenient to carry
all data in a single file, MsWord also supports the "MHTML" format (or
"MHT" for short), i.e. an encapsulation of a whole HTML tree into a
single file encoded in MIME multipart format. This format can be
generated interactively from MsWord by calling the "SaveAs" menu and
choosing the F<.mht> extension.

Documents saved with a F<.mht> extension will not directly 
reopen in MsWord : when clicking on such documents, Windows
chooses Internet Explorer as the default display program.
However, these documents can be simply renamed with a
F<.doc> extension, and will then open directly in MsWord.
The same can be done with WordML or RTF documents.
That is to say, MsWord is able to recognize the internal
format of a file, without any dependency on the 
filename.

C<MsOffice::Word::HTML::Writer> helps you to programatically generate
MsWord documents in MHT format. The advantage of this technique is
that one can rely on standard HTML mechanisms for layout control, such
as styles, tables, divs, etc. Of course this markup can be produced
using your favorite HTML module; the added value
of C<MsOffice::Word::HTML::Writer> is to help building the 
MIME multipart file, and provide some abstractions for 
representing MsWord-specific features (headers, footers, fields, etc.).


The MHT format is probably the most convenient
way for programmatic document generation, because

=over

=item *

unlike Excel, MsWord native binary format is unpublished and therefore
cannot be generated without the MsWord executable.

=item *

remote control of the MsWord program through an OLE connection,
as in L<Win32::Word::Writer|Win32::Word::Writer>, requires a
local installation of Microsoft Office, and is not well
suited for servers because the MsWord program might hang
or might open dialog boxes that require user input.

=item *

generation of documents in RTF is possible, but 
requires deep knowledge of the RTF structure
--- see L<RTF::Writer>.

=item *

generation of documents in "WordML" also requires
deep knowledge of WordML structure.

=back

By contrast, C<MsOffice::Word::HTML::Writer> allows you to 
produce documents even with little knowledge of MsWord.
One word of warning, however : opening MHT documents in MsWord is
slower than native binary or RTF documents, because MsWord needs to
parse the HTML, compute the layout and convert it into its internal
representation.  Therefore MHT format is not recommended for large
documents.

B<Note> : this first release of C<MsOffice::Word::HTML::Writer>
is still in an exploratory phase; the programming interface
may change in future versions.

=head1 METHODS

=head2 new

    my $doc = MsOffice::Word::HTML::Writer->new(%params);

Creates a new writer object. Optional parameters are :

=over

=item title

document title

=item styles

a string containing CSS style definitions for the main document

=item hf_styles

a string containing CSS style definitions for headers and footers

=back

=head2 write

  $doc->write("<p>hello, world</p>");

Adds some HTML into the document body.

=head2 break_page

Inserts a page break.

=head2 break_section

    $doc->break_section(%params);

Opens a new section within the document.  Each section in a document
can have independent pagination (margins, columns, headers and
footers, etc.). Parameters are :

=over

=item header

Header content (in html)

=item first_header

Header for the first page of that section.

=item footer

Footer content (in html).

=item first_footer

Footer for the first page.

=item new_page

If true, a page break will be inserted before the new section.

=back

Following calls to the C<write> method will add content to
the newly created section.


=head2 save_as

  $doc->save_as("C:/Temp/mydoc");

Generates the MIME document and saves it at the given location.
If no extension is present, file extension F<.doc> will be added 
by default.


=head2 content

Returns the whole MIME-encoded document as a single string.
Useful if you don't want to save the document into a file, but 
want to do something else like embedding it in a message
or a ZIP file, or returning it as an HTTP response.


=head2 field

  my $html = $doc->field($fieldname, $args); 

Returns HTML markup for a MsWord field.
Optional C<$args> is a string with arguments or flags for
the field. See MsWord help documentation for the list of 
field names and their associated arguments or flags. 
Here are some examples :

  my $header = sprintf "%s of %s", $doc->field('PAGE'), 
                                   $doc->field('NUMPAGES');
  my $footer = sprintf "created at %s, printed at %s", 
                 doc->field(CREATEDATE => '\\@ "d MM yyyy"'),
                 doc->field(PRINTDATE  => '\\@ "dddd d MMMM yyyy" \\* Upper');

=head2 attach

  $doc->attach($localname, $filename);
  $doc->attach($localname, "<", \$content);
  $doc->attach($localname, "<&", $filehandle);

Adds an attachment into the document; the attachment will be encoded
as a MIME part and will be accessible under C<files/$localname>.

The remaining arguments to C<attach> specify the source of the attachment;
they are directly passed to L<perlfunc/open> and therefore have the same
API flexibility : you can specify a filename, a reference to a memory
variable, a reference to another filehandle, etc.


=head1 TO DO

This is still exploratory; many features need to be added. 
For example:

  - headers/footers for odd/even pages
  - link same header/footers across several sections
  - section formatting (margins, page size, columns, etc)




=head1 AUTHOR

Laurent Dami, C<< <laurent DOT dami AT etat DOT geneve DOT ch> >>

=head1 BUGS

Please report any bugs or feature requests to
C<bug-win32-word-html-writer at rt.cpan.org>, or through the web interface at
L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=MsOffice-Word-HTML-Writer>.
I will be notified, and then you'll automatically be notified of progress on
your bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc MsOffice::Word::HTML::Writer

You can also look for information at:

=over 4

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/MsOffice-Word-HTML-Writer>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/MsOffice-Word-HTML-Writer>

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=MsOffice-Word-HTML-Writer>

=item * Search CPAN

L<http://search.cpan.org/dist/MsOffice-Word-HTML-Writer>

=back

=head1 SEE ALSO

L<Win32::Word::Writer>, L<RTF::Writer>, L<Spreadsheet::WriteExcel>.


=head1 COPYRIGHT & LICENSE

Copyright 2008 Laurent Dami, all rights reserved.

This program is free software; you can redistribute it and/or modify it
under the same terms as Perl itself.

=cut

1; # End of MsOffice::Word::HTML::Writer
