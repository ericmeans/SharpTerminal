Features:
1. View communication as either ASCII text or hexadecimal values
2. Save full session transcripts in multiple formats for easy analysis
3. Open previous session transcripts
4. Easy entry of binary data (prepend with 0x for hex entry)
5. Unlimited command history (up and down arrow in Send box)
6. Show or hide connection, control line, etc events.
7. Colored text for ease of distinguishing between sent and received data
8. Multithreaded for responsiveness
9. Prettier than Hyperterminal

Future Enhancements:
1. Error Handling is not completely up to snuff.  It won't crash, but it's not as pretty as it could be.
2. Allow user to select encoding for "Text" mode.
3. Enable DTR handshaking
4. Consider allowing the intermixing of ASCII and binary?

System Requirements:
1. Microsoft .Net Framework, version 1.1.  It might run against 1.0, I haven't tried.
2. Internet Explorer (any version 4 or later should work AFAIK).
3. One or more serial ports.

Licensing:
SharpTerminal use is not limited; you may copy it, redistribute it freely, use it in a business, install it on a rocket and shoot it to the moon, or anything else I haven't mentioned here, with the following restrictions:
1. No claiming it is your own work.  You must include this ReadMe.txt file, UNMODIFIED, any time you redistribute it.
2. Actually, that's pretty much it.  If you really need specifics, see http://creativecommons.org/licenses/by/1.0/

Questions? Comments? Bugs? Feature Requests?
Visit http://www.randomtree.org/sharpterminal/ or e-mail code@randomtree.org

All code, text, and images copyright 2004 Eric Means.

ChangeLog:
2004-06-25: Fixed issues with COM ports > COM4 not opening.
			Enabled Windows XP Visual Styles.