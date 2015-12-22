# <@LICENSE>
# Licensed to the Apache Software Foundation (ASF) under one or more
# contributor license agreements.  See the NOTICE file distributed with
# this work for additional information regarding copyright ownership.
# The ASF licenses this file to you under the Apache License, Version 2.0
# (the "License"); you may not use this file except in compliance with
# the License.  You may obtain a copy of the License at:
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# </@LICENSE>

=head1 NAME

OLE2Macro - Look for Macro Embedded Microsoft Word and Excel Documents

=head1 SYNOPSIS

  loadplugin     ole2macro.pm
  body MICROSOFT_OLE2MACRO eval:check_microsoft_ole2macro()
  score MICROSOFT_OLE2MACRO 3
  
=head1 DESCRIPTION

Detects embedded OLE2 Macros embedded in Word and Excel Documents. Based on:
https://blog.rootshell.be/2015/01/08/searching-for-microsoft-office-files-containing-macro/

10/12/2015 - Jonathan Thorpe - jthorpe@conexim.com.au

=back

=cut

package OLE2Macro;

use Mail::SpamAssassin::Plugin;
use Mail::SpamAssassin::Util;
use strict;
use warnings;
use bytes;
use re 'taint';

use vars qw(@ISA);
@ISA = qw(Mail::SpamAssassin::Plugin);

# constructor: register the eval rule
sub new {
  my $class = shift;
  my $mailsaobject = shift;

  # some boilerplate...
  $class = ref($class) || $class;
  my $self = $class->SUPER::new($mailsaobject);
  bless ($self, $class);

  $self->register_eval_rule("check_microsoft_ole2macro");

  return $self;
}

sub check_microsoft_ole2macro {
  my ($self, $pms) = @_;

  _check_attachments(@_) unless exists $pms->{nomacro_microsoft_ole2macro};

  return $pms->{nomacro_microsoft_ole2macro};
}

sub _check_attachments {
  my ($self, $pms) = @_;
  my $marker1 = "\xd0\xcf\x11\xe0"; 
  my $marker2 = "\x00\x41\x74\x74\x72\x69\x62\x75\x74\x00"; 

  $pms->{nomacro_microsoft_ole2macro} = 0;

  foreach my $p ($pms->{msg}->find_parts(qr/./, 1)) {
    my ($ctype, $boundary, $charset, $name) =
      Mail::SpamAssassin::Util::parse_content_type($p->get_header('content-type'));

    my $cte = lc($p->get_header('content-transfer-encoding') || '');
    $ctype = lc $ctype;

    if ($cte =~ /base64/){
       my $contents = $p->decode();
       if (index($contents, $marker1) > -1 && 
           index($contents, $marker2) > -1) {
          $pms->{nomacro_microsoft_ole2macro} = 1;
       }
    }
  }
}

1;
