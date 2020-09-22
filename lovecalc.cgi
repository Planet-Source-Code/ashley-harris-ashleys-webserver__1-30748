#! /usr/bin/perl

print "Content-type: text/html\n\n";
print "<HTML><HEAD><TITLE>Ashley's Love Calculator</TITLE></HEAD><BODY bgcolor=FF6666>\n";

$req = $ENV{QUERY_STRING};
#$req = "male=&female=";

if (length($req) > 0)
{
 %in = ();
 @vals = split(/\&/,$req);
 foreach $a (@vals)
 {
  $a =~ m/([a-zA-Z0-9\%\+]{1,40})\=([a-zA-Z0-9\%\+]{1,40})/;
  if ($1 eq "") {next}
  ($key,$value) = ($1,$2);
  $key =~ s/\+/ /g;
  $value =~ s/\+/ /g;
  $key =~ s/%([A-Fa-f0-9]{2})/pack("c",hex($1))/ge;
  $value =~ s/%([A-Fa-f0-9]{2})/pack("c",hex($1))/ge;
  $in{$key} = uc($value);
 }
 @ms = split(//,"ASHLEYSWEBSERVER");
 $male = $in{"male"};
 $female = $in{"female"};
 
 if ($male eq "") {$male = "YOUR AVERAGE GUY"}
 if ($female eq "") {$female = "YOUR AVERAGE GIRL"}
 
 my $c = 0;
 for $a (@ms)
 {
  $c = $c + ct($male,$a);
  $c = $c + ct($female,$a);
 }
 $total = $c /(length($male)+length($female))*50;
 print "Doctor Love says that the chances of <B>$male</B> and <B>$female</B>'s relationship being successfull is $total%<P>\n";
}
print "Ask Dr Love about the success of what relationship? (Fill out one or both fields)<BR>\n";
print "<FORM method=GET action=\"lovecalc.cgi\">\n";
print "<TABLE>\n";
print "<tr><td>Male: </td><td><INPUT type=text name=male></td></tr>\n";
print "<tr><td>Female: </td><td><INPUT type=text name=female></td></tr></table><P>\n";
print "<INPUT type=submit value=\"Submit to Dr Love\">\n";
print "</BODY></HTML>";
  
sub ct
{
 my $string = shift;
 my $what = shift;
 return ($string =~ s/$what/$what/g);
}
