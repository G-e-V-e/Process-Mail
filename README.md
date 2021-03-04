# Process-Mail
Do everything needed to manage, build, process and send html formatted mail messages based on just 1 variable and 1 PowerShell function.

"Everything happens for a reason", Aristotle believed. That may be a wrong assumption given the current state of philosophy and science, but at least sending mails should.

So, the basis for this Powershell function is to define at least one and preferrably a lot more reasons why you would want - or need - mails to email addresses in your company or to the outside world.

Say you want to communicate that a given process succeeded, you could call that reason 'OK' and send it to everyone. On the other hand, a 'NotOK' process may be even more reason to send a mail, probably just to yourself as the developper. Extend that to Create, Update and Remove events and you may already have 6 different reasons:

CreateOK, CreateNotOK, UpdateOK, UpdateNotOK, RemoveOK and RemoveNotOK

The names and number of reasons are almost entirely limited by your imagination, almost because 'List' is reserved for a specific meaning.

What do you need more to send mail messages?

* A mail sender, the name figuring as the sender of the mail message. Since the mail messages are sent automatically from a script, not manually from a mailbox, you are free to supply any well formed email-address. Perhaps it might be your own, existing, personal email-address or a dummy one such as 'noreply.scriptname@my.company.com'. Don't supply a completely random chosen sender name because that could get you in trouble with your security department...
* One or more mail servers to connect and hand over your mail to. A mail server knows where to forward that mail message, to have it delivered at its destination. Big companies usually have at least 2 mailservers and a bridge (loadbalancer) between them. Set the bridge server first, then the real mail servers. If for some reason a mail server refuses connection, the mail handover is retried 3 times with 100 millisecond intervals. After that, the next mail server is tried to contact 3 times and so on. Only after the last server refused to connect that the mail message is considered undeliverable. 
* One or more destination email addresses of persons who have any intrest in receiving these mail messages.
* Not every defined mail destination receives all mails being sent: you or preferrably they can choose for which reason(s) they will receive mail messages.
* Mail messages are html formatted. While there are builtin formatting defaults for every html element used, these defaults can be overridden globally and personally. A visually challenged person could receive mails with a larger font size, someone with colorblindness issues could receive mails in contrasting colors to his/her liking.
* The body of the mail message consists of what your want it to. You can supply text, variable contents and/or objects. A mail for a failed process will probably include the errormessage variable or object $error[0].
* The body can be formatted in 3 ways: as text (formatting option 'Text'), as detailed [type]value information (option 'Full') and as something in between ('Basic'). For debugging purposes, option 'Full' is recommended but not everybody needs or understands that level of detail.
* Finally, there is the 'List' option. People who are only interested in getting an overview of what happened during execution rather than receiving a new mail message at every occurrance of reason, can receive an accumulated list by any number of reasons. In 'Lists', you define for which reasons you want to collect a list. When all the processing is done, you call this PowerShell function with reason 'List' and those who have 'List' specified as format for a reason will get that reason's overview sent to them, at least if there occurred reason events during execution. You can html format that list output as 'Text', 'Basic' or 'Full' as requested by the destination person.

All the above information items are held in one (1) single hashtable variable, its default name being $Mail. It has to be accessible from function calls, so better define $Mail in the Global or Script scope. Fully used, the $Mail variable contains 5 hashtable levels. While that $Mail variable can be constructed from hardcoded information in the calling script, by far the easiest way to populate the information required in that $Mail variable is to create it based on the contents of an IniFile.

An IniFile example is given as a starting point to understand what every information item is, how to specify it and how it relates to other items, and from there build your own IniFile. You may find my repository 'Not-Just-Another-Reincarnation-Of-IniFile' useful because it is capable of generating any multilevel hashtable in just one go.

One last thing: if you send a mail message without supplying a reason, that message will probably reach me. You could use that opportunity to ask questions, so don't include company secrets but some way to contact you would be useful...

geve.one2one@gmail.com
