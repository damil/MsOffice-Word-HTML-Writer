use utf8;
use strict;
use warnings;
use lib "../lib";

use MsOffice::Word::HTML::Writer;


my $hf_styles = <<"";
.pj_footer {margin:0; padding: 0}
.tech_size {font-size: 5.0pt}
TD         {padding: 0cm 0.12cm 0cm 0.12cm }

my $footer = <<__END_FOOTER__;
<p class="pj_footer" style="font-size: 9.0pt">
  <b>[% doc.pied_page_1 %]</b>
  [% doc.pied_page.buffer_ %]
</p>

<p class="pj_footer" style="font-size: 9.0pt; text-align: center">
  <b>[% doc.pied_page_2 %][% doc.accuse_reception %]</b>
</p>

<table border=0 cellspacing=0 cellpadding=0 
       style='border-collapse:collapse;width:16.62cm'>


 <tr height='0.35cm'>
  <td colspan="2">
   <table border=0 cellspacing=0 cellpadding=0 
          style='border-collapse:collapse;width:16.62cm'>
    <tr>
       <td width="1.62cm" valign=bottom 
           style='border:none;border-top:solid black 1.0pt; font-size: 5.0pt'>
         [% doc.tech.NOM_DOCUMENT %]
       </td>

       <td width="14cm" valign=bottom 
           style='border:none;border-top:solid black 1.0pt; font-size: 8.0pt; text-align: center'>
             [% doc.EN_TETE_GLOBAL %]
             [% doc.exp.TPHONE %]
             [% doc.exp.FAX %]
       </td>

       <td width="1cm" valign=bottom 
           style='border:none;border-top:solid black 1.0pt; font-size: 5.0pt'>
         [% doc.tech.OPER_MODIF %]
       </td>
    </tr>
  </table>
 </td>
 </tr>

 <tr height='0.2cm'>
  <td width="50%" style="font-size: 5.0pt">
    [% doc.tech.MODE_EXP %]
  </td>
  <td width="50%" style="font-size: 5.0pt">
    [% doc.tech.TEST %]
  </td>
 </tr>

 <tr height='0.3cm'>
  <td width="50%" style="font-size: 8pt">[% doc.tech.INFO %]</td>
  <td width="50%" style="font-size: 8pt">[% doc.tech.INFO %]</td>
 </tr>

 <tr height='0.3cm'>
  <td width="50%" style="font-size: 8pt"></td>
  <td width="50%" style="font-size: 8pt">[% doc.tech.INFO %]</td>
 </tr>

 <tr height='0.3cm'>
  <td colspan="2" style="font-size: 8pt">
    <b>[% doc.tech.FLAG_HANDICAPE %]</b>
  </td>
 </tr>
</table>
__END_FOOTER__


my $doc = MsOffice::Word::HTML::Writer->new(
  title     => "My new doc",
  # hf_styles => $hf_styles,
 );



$doc->create_section(
  footer    => $footer,
);


$doc->write(<<"");
   <P>HELLO, WORLD</P>
   <P>THIS IS A GENERATED DOC</P>
   <p>il était une bergère</p>

$doc->save_as('pj.doc');



__END__

--HEADER_TEMPLATE---

<p class=MsoHeader><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p>&nbsp;</o:p></span></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=704
 style='width:528.0pt;margin-left:-35.45pt;border-collapse:collapse;mso-padding-alt:
 0cm 3.5pt 0cm 3.5pt'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;
  height:45.4pt'>
  <td width=51 colspan=3 rowspan=2 valign=top style='width:38.15pt;padding:
  0cm 3.5pt 0cm 3.5pt;height:45.4pt'>
  <p class=MsoHeader><!--[if gte vml 1]><v:shapetype id="_x0000_t75"
   coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe"
   filled="f" stroked="f">
   <v:stroke joinstyle="miter"/>
   <v:formulas>
    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
    <v:f eqn="sum @0 1 0"/>
    <v:f eqn="sum 0 0 @1"/>
    <v:f eqn="prod @2 1 2"/>
    <v:f eqn="prod @3 21600 pixelWidth"/>
    <v:f eqn="prod @3 21600 pixelHeight"/>
    <v:f eqn="sum @0 0 1"/>
    <v:f eqn="prod @6 1 2"/>
    <v:f eqn="prod @7 21600 pixelWidth"/>
    <v:f eqn="sum @8 21600 0"/>
    <v:f eqn="prod @7 21600 pixelHeight"/>
    <v:f eqn="sum @10 21600 0"/>
   </v:formulas>
   <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
   <o:lock v:ext="edit" aspectratio="t"/>
  </v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style='width:28.5pt;
   height:47.25pt' fillcolor="window">
   <v:imagedata src="image001.png" o:title=""/>
  </v:shape><![endif]--></p>
  </td>
  <td width=653 colspan=5 valign=top style='width:489.85pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:45.4pt'>
  <p class=MsoHeader style='line-height:8.0pt;mso-line-height-rule:exactly'><span
  lang=FR-CH style='mso-ansi-language:FR-CH'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='font-size:11.0pt;mso-bidi-font-size:10.0pt'>R&eacute;publique
  et canton de Gen&egrave;ve<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span>Gen&egrave;ve, <a name="M_00000000000000000000000008513">&lt;PAL_D_ACT_DATE1_CLAIR&gt;</a></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004496"><b
  style='mso-bidi-font-weight:normal'><span style='font-size:11.0pt;mso-bidi-font-size:
  10.0pt'>&lt;PAL_D_ATT_C_JUR&gt;</span></b></a></p>
  <p class=MsoNormal><a name="M_00000000000000000000000012967"><b
  style='mso-bidi-font-weight:normal'>&lt;PAL_D_ATT_EN_TETE_1&gt;</b></a><b
  style='mso-bidi-font-weight:normal'><o:p></o:p></b></p>
  <p class=MsoNormal><a name="M_00000000000000000000000012968"><b
  style='mso-bidi-font-weight:normal'>&lt;PAL_D_ATT_EN_TETE_2&gt;</b></a><b
  style='mso-bidi-font-weight:normal'><o:p></o:p></b></p>
  <p class=MsoNormal><a name="M_00000000000000000000000012969"><b
  style='mso-bidi-font-weight:normal'>&lt;PAL_D_ATT_EN_TETE_3&gt;</b></a><b
  style='mso-bidi-font-weight:normal'><o:p></o:p></b></p>
  <p class=MsoNormal><a name="M_00000000000000000000000012970"><b
  style='mso-bidi-font-weight:normal'>&lt;PAL_D_ATT_EN_TETE_4&gt;</b></a><b
  style='mso-bidi-font-weight:normal'><o:p></o:p></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;page-break-inside:avoid;height:8.5pt'>
  <td width=653 colspan=5 valign=top style='width:489.85pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.5pt'>
  <p class=MsoHeader align=right style='text-align:right;line-height:8.0pt;
  mso-line-height-rule:exactly'><a name="M_00000000000000000000000004505"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt;layout-grid-mode:line'>&lt;PAL_D_PRO_N_EXT_PROC&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt;layout-grid-mode:line'> <a
  name="M_00000000000000000000000004506">&lt;PAL_D_ATT_CHAMBRE&gt;</a> <a
  name="M_00000000000000000000000004507">&lt;PAL_D_ATT_USER_ID&gt;</a> <a
  name="M_00000000000000000000000004905">&lt;PAL_D_ATT_C_TYPE_ATTR&gt;</a> <a
  name="M_00000000000000000000000013100">&lt;PAL_D_AVO_N_CASE&gt;</a></span><span
  lang=FR-CH style='mso-ansi-language:FR-CH'><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2;page-break-inside:avoid;height:12.0pt;mso-row-margin-left:
  17.15pt'>
  <td style='mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm'
  width=23><p class='MsoNormal'>&nbsp;</td>
  <td width=369 colspan=5 rowspan=3 valign=top style='width:277.05pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:12.0pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004498"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&lt;PAL_D_ATT_ADR_JUR_1&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004499"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&lt;PAL_D_ATT_ADR_JUR_2&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004500"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&lt;PAL_D_ATT_ADR_JUR_3&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004501"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&lt;PAL_D_ATT_ADR_JUR_4&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004502"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&lt;PAL_D_ATT_ADR_JUR_5&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>
  <p class=MsoHeader><a name="M_00000000000000000000000012623"><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&lt;PAL_D_ATT_ADR_JUR_6&gt;</span></a><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>
  </td>
  <td width=246 valign=top style='width:184.2pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:12.0pt'>
  <p class=MsoNormal style='tab-stops:5.0cm'><a
  name="M_00000000000000000000000011200"><b style='mso-bidi-font-weight:normal'><span
  style='layout-grid-mode:line'>&lt;PAL_D_DES_TYPE_ENVOI&gt;</span></b></a><span
  style='layout-grid-mode:line'><o:p></o:p></span></p>
  </td>
  <td width=66 valign=top style='width:49.6pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:12.0pt'>
  <p class=MsoNormal align=right style='text-align:right;tab-stops:5.0cm'><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt;layout-grid-mode:line'><o:p>&nbsp;</o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:3;page-break-inside:avoid;height:11.35pt;mso-row-margin-left:
  17.15pt'>
  <td style='mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm'
  width=23><p class='MsoNormal'>&nbsp;</td>
  <td width=246 valign=top style='width:184.2pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:11.35pt'>
  <p class=MsoNormal style='tab-stops:5.0cm'><a
  name="M_00000000000000000000000004503"><span style='layout-grid-mode:line'>&lt;PAL_D_DES_I_ELEC_DOM&gt;</span></a><span
  style='layout-grid-mode:line'><o:p></o:p></span></p>
  </td>
  <td width=66 valign=top style='width:49.6pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:11.35pt'>
  <p class=MsoNormal align=right style='text-align:right;tab-stops:5.0cm'><span
  style='font-size:8.0pt;mso-bidi-font-size:10.0pt;layout-grid-mode:line'><o:p>&nbsp;</o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:4;page-break-inside:avoid;height:66.6pt;mso-row-margin-left:
  17.15pt'>
  <td style='mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm'
  width=23><p class='MsoNormal'>&nbsp;</td>
  <td width=312 colspan=2 valign=bottom style='width:233.8pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:66.6pt'>
  <p class=MsoNormal style='tab-stops:5.0cm'><a
  name="M_00000000000000000000000011199"><span style='layout-grid-mode:line;
  mso-bidi-font-weight:bold'>&lt;PAL_D_DES_BUFFER_ADRESSE&gt;</span></a><span
  style='layout-grid-mode:line;mso-bidi-font-weight:bold'><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5;page-break-inside:avoid;mso-row-margin-left:17.15pt'>
  <td style='mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm'
  width=23><p class='MsoNormal'>&nbsp;</td>
  <td width=19 valign=top style='width:14.05pt;padding:0cm 3.5pt 0cm 3.5pt'>
  <p class=MsoNormal><o:p>&nbsp;</o:p></p>
  </td>
  <td width=39 colspan=2 valign=top style='width:29.05pt;border-top:solid windowtext 1.0pt;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:none;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  background:#F2F2F2;mso-shading:white;mso-pattern:gray-5 auto;padding:0cm 3.5pt 0cm 3.5pt'>
  <p class=MsoNormal>R&eacute;f&nbsp;:</p>
  </td>
  <td width=210 valign=top style='width:157.8pt;border-top:solid windowtext 1.0pt;
  border-left:none;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  background:#F2F2F2;mso-shading:white;mso-pattern:gray-5 auto;padding:0cm 3.5pt 0cm 3.5pt'>
  <p class=MsoNormal><a name="M_00000000000000000000000004516"><b
  style='mso-bidi-font-weight:normal'>&lt;PAL_D_PRO_N_EXT_PROC&gt;</b></a> <a
  name="M_00000000000000000000000004517">&lt;PAL_D_ATT_CHAMBRE&gt;</a> <a
  name="M_00000000000000000000000004518">&lt;PAL_D_ATT_USER_ID&gt;</a> <a
  name="M_00000000000000000000000004519">&lt;PAL_D_ATT_C_TYPE_ATTR&gt;</a></p>
  <p class=MsoNormal><a name="M_00000000000000000000000004545">&lt;PAL_D_ACT_N_EXT_DECIS&gt;</a></p>
  </td>
  <td width=102 valign=top style='width:76.15pt;padding:0cm 3.5pt 0cm 3.5pt'>
  <p class=MsoNormal><o:p>&nbsp;</o:p></p>
  </td>
  <td width=312 colspan=2 valign=top style='width:233.8pt;padding:0cm 3.5pt 0cm 3.5pt'>
  <p class=MsoNormal><span style='layout-grid-mode:line'><o:p>&nbsp;</o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:6;mso-yfti-lastrow:yes;page-break-inside:avoid;
  height:7.8pt;mso-row-margin-left:17.15pt'>
  <td style='mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm'
  width=23><p class='MsoNormal'>&nbsp;</td>
  <td width=19 valign=top style='width:14.05pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.8pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=39 colspan=2 valign=top style='width:29.05pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;
  border-right:none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
  solid windowtext .5pt;background:#F2F2F2;mso-shading:white;mso-pattern:gray-5 auto;
  padding:0cm 3.5pt 0cm 3.5pt;height:7.8pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=210 valign=top style='width:157.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  background:#F2F2F2;mso-shading:white;mso-pattern:gray-5 auto;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.8pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>&agrave;
  rappeler lors de toute communication<o:p></o:p></span></p>
  </td>
  <td width=102 valign=top style='width:76.15pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.8pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=312 colspan=2 valign=top style='width:233.8pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.8pt'>
  <p class=MsoNormal><span style='layout-grid-mode:line'><o:p>&nbsp;</o:p></span></p>
  </td>
 </tr>
 <![if !supportMisalignedColumns]>
 <tr height=0>
  <td width=23 style='border:none'></td>
  <td width=19 style='border:none'></td>
  <td width=9 style='border:none'></td>
  <td width=29 style='border:none'></td>
  <td width=210 style='border:none'></td>
  <td width=102 style='border:none'></td>
  <td width=246 style='border:none'></td>
  <td width=66 style='border:none'></td>
 </tr>
 <![endif]>
</table>

<p class=MsoNormal><span style='layout-grid-mode:line'><o:p>&nbsp;</o:p></span></p>

<p class=MsoHeader><o:p>&nbsp;</o:p></p>




--FOOTER_TEMPLATE---
.font_9pt {font-size: 9.0pt}
.font_8pt {font-size: 8.0pt}
.centered  {text-align: center}
.tech_size {font-size: 5.0pt}
TD         {padding: 0cm 0.12cm 0cm 0.12cm }


<p class="font_9pt">
  <b>[% doc.pied_page_1 %]</b>
  [% doc.pied_page.buffer_ %]
</p>

<p class="font_9pt centered">
  <b>[% doc.pied_page_2 %]</b>
  <b>[% doc.accuse_reception %]</b>
</p>

<table border=0 cellspacing=0 cellpadding=0 width="16.62cm"
       style='border-collapse:collapse'>

 <tr height='0.35cm'>

  <td width="1.62cm" valign=bottom 
      style='border:none;border-top:solid black 1.0pt'>
    <p class="tech_size">[% doc.tech.NOM_DOCUMENT %]</p>
  </td>

  <td width="14cm" valign=bottom 
      style='border:none;border-top:solid black 1.0pt'>
    <p class="font_8pt centered">
      [% doc.EN_TETE_GLOBAL %]
      [% doc.exp.TPHONE %]
      [% doc.exp.FAX %]
    </p>
  </td>

  <td width="1cm" valign=bottom 
      style='border:none;border-top:solid black 1.0pt'>
    <p class="tech_size">[% doc.tech.OPER_MODIF %]</p>
  </td>
 </tr>
</table>

<table border=0 cellspacing=0 cellpadding=0 width="16.62cm"
       style='border-collapse:collapse'>
 <tr height='0.2cm'>
  <td width="50%">
    <p class="tech_size">[% doc.tech.MODE_EXP %]</p>
  </td>
  <td width="50%">
    <p class="tech_size">[% doc.tech.TEST %]</p>
  </td>
 </tr>

 <tr height='0.3cm'>
  <td width="50%">
    <p class="font_8pt">[% doc.tech.INFO %]</p>
  </td>
  <td width="50%">
    <p class="font_8pt">[% doc.tech.INFO %]</p>
  </td>
 </tr>

 <tr height='0.3cm'>
  <td width="50%">
  </td>
  <td width="50%">
    <p class="font_8pt">[% doc.tech.INFO %]</p>
  </td>
 </tr>

 <tr height='0.3cm'>
  <td colspan="2">
    <p class="font_8pt"><b>[% doc.tech.FLAG_HANDICAPE %]</b></p>
  </td>
 </tr>
</table>






