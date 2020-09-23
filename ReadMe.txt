Lan Chat

Thanks to:
	J.-A.Mock - the beastily .bas that provides my application with the transparency ability.

Any Requests?

Yeah, vote for me, I'm only 16 and I cant afford fancy software (well legally any way ;)) to develop nice applications! So, do me a big favour, vote and support the code I keep churning out! Cheers!

Why was it made?

1. I was bored.
2. Where I am a student, all the machines run Win 2K, so net send is quite tempting but logged (doh!) so if you want to send a cheeky message about a lecturer you have to be careful. Also, winchat.exe is not pretty and I though I'd make my own that start up when you log on.

Purpose?

Simple UDP (user datagram protocol) chat program that allows a conversation to be staged across a network with ease. There are only clients invloved, i.e, it is a peer to peer based communication system.

What does it do?

When you start up LanChat.exe, two Winsock Control (MSWinsck.ocx is included) are initalized and set to UDP. One binds itself to 1200 and accepts any incoming connections. The second WinSock connects to the broadcast address (255.255.255.255) which should work on any network in theory, if you have difficulties, try changiing this to a more localasied address such as xxx.xxx.xxx.0.

When data is sent, it is broadcast to all machines, and the ones that are listening will stick the data in the txtmsgs text box. When a user sends data, what the user has sent is actaully put in txtmsgs by socket bound to port 1200, so an internal communcation is going on, this was the easiest way of getting it to work, but not probably not that efficient, but it works none the less.

Like other program I have written, it will start up when the machine is logged on as it writes a regisrty key that causes it do so.

Under Windows 2000 and XP, (not sure about NT4) you can set the transparency of the form with the slider at the bottom of the form, this prevents spying lectures (in my case!) or system administrators seeing how you are wasting the network bandwidth from a distance, and also it looks good! ;) This functionallity was possible thanks to the example on PSC by J.-A.Mock (cheers for the example, very useful).

v2.0

Whats new?

There are now support for commands, which enable a user to perform various tasks;

1. /:listusers

Quite obvious - justs lists all the users logged on

2. /:kickuser [Windows User Name]

Useful if you want to ban somebody temporarily. Windows user name should be used instead of screen name. The ban period is 10 minutes which stored in the registry. The user cannot get LaNcHaT to run until this time has lapsed.

Commands can be typed secretely by left clicking the message text box whilst holding the shift key - handy to stop others from seeing the commands and using them on you!

Flasher!

After a few tests, I have finally got the form to flash when a message is received and user focus is elsewhere - for some reason GotFocus and LostFocus do not seem to work on Win2k with VB6, so whether the form flashes depends on the WindowState property, and stops flashing when the user begins to type a message.

Enjoy!