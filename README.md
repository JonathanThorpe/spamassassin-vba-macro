# SpamAssassin Plugin for detecting VBA/OLE2 Macros in Word/Excel Documents

SpamAssassin plugin for detecting VBA macros in Microsoft Office Documents.

## Motivation ##

This plugin was built in response to increasing phishing/malware found in Microsoft Word and Excel Documents that seems to get past the majority of Virus Scanners when new waves first start.

Most installations running SpamAssassin also run ClamAV and while ClamAV has a configuration option to block OLE2Macros (OLE2BlockMacros), this doesn't work well in practice because surprisingly, some people still send emails with legitimate VBA Macros.

This plugin is designed to crudely (but so far, effectively) identify VBA/OLE2 Macros embedded in Microsoft Word and Excel Documents and allows you to assign a configurable score to these documents such that other factors can contribute to an overall score before considering the email spam.



