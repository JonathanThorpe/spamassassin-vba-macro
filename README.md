# SpamAssassin Plugin for detecting VBA/OLE2 Macros in Word/Excel Documents

SpamAssassin plugin for detecting VBA macros in Microsoft Office Documents.

## Motivation ##

This plugin was built in response to increasing phishing/malware found in Microsoft Word and Excel Documents that seems to get past the majority of Virus Scanners when new waves first start.

Most installations running SpamAssassin also run ClamAV and while ClamAV has a configuration option to block OLE2Macros (OLE2BlockMacros), this doesn't work well in practice because surprisingly, some people still send emails with legitimate VBA Macros.

This plugin is designed to crudely (but so far, effectively) identify VBA/OLE2 Macros embedded in Microsoft Word and Excel Documents by looking at a signature in the file only present when VBA Macros are used and allows you to assign a configurable score to these documents such that other factors can contribute to an overall score before considering the email spam.

## Installation ##

Simply place the following in your SpamAssassin configuration file along with the ole2macro.pm in a directory where SpamAssassin is able to find it (i.e. /etc/spamassassin). For instance ```/etc/spamassassin/conf.d/90_ole2macro.cf```

```
loadplugin OLE2Macro ole2macro.pm

body MICROSOFT_OLE2MACRO eval:check_microsoft_ole2macro()
describe MICROSOFT_OLE2MACRO Has an attachment that contains an OLE2 Macro
score MICROSOFT_OLE2MACRO 4
```

In most cases, you will want to adjust the score accordingly, however, with a cut-off score of 5.5, 4.0 as described above has worked well in a production environment with diverse users.

## To Do ##

In isolated cases, Word/Excel Documents are contained within Zip archives which are undetected by this software. Investigate unpacking zip files and examining such files contained within.
