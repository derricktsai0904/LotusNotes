Okay, this one's not really encryption in the sense of securely transforming data into something illegible, but it's a good place to start. Base64 (described in RFC 2045, among others) is a "classic" way of encoding binary data into text strings, and it's been used for transferring e-mail attachments for years. It's also used in conjunction with things like basic HTTP authentication and PGP signatures. The ability to encode and decode Base64 data is a good thing to have in your toolbox, because you'll run into it in lots of places.

Here's a LotusScript implementation of how to encode and decode Base64 strings. I haven't tested it recently , but it should work in many (all?) of the 4.x versions of Notes as well as R5 and R6 -- and heck, it ought to work in Visual Basic, too. (Udpated December 28: fixed TrimBytesFromFile function)

[程式來源 Base64v14.lss]:https://github.com/derricktsai0904/LotusNotes/blob/master/Base64%20Encoding/Base64v14.lss "Base64v14.lss"
[程式來源 Base64v14.lss]
