# VersionParser

## The problem
Excel can't sort a table based on a version number in the format #.#.#(.#.#.#...etc) properly, as this isn't a recognised format. The closest it can do is sort the values bsed on their string values, meaning that 1.2.11 is interpreted as being less than 1.2.2.

## The solution
Write a bit of VBA that parses the version number string into a floating point number.

In fact, it's (slightly) cleverer than it sounds, as simply parsing 1.2.3 into 1.23 wouldn't work with any revision numbers greater than 10 (as 1.2.11 -> 1.31 rather than 1.211). Additionally, even if it did parse into 1.211, it would be impossible to tell whether the original version was 1.2.11 or 1.2.1.1

As hinted by the previous sentence, VersionParser (or VER2NUM as the function is labelled) can also cope with versioning levels beyond the typical Major.Minor.Patch format.

## Limitations
Aside from the bad formatting/layout/practices that I've presumably used (It's my first bit of VBA... give me a break!), VersionParser won't be able to cope with any version numbers greater than 1 million. If, for whatever bizarre reason, you need to increase this limit, it's relatively simple to do - just change the 5 on line 17 to a 6 or a 7 for flexibility of over 10 million revisions!

Also, yes, I realise I really ought to comment more.

## Contributions
Feel free to improve this project in any way by simply submitting a Pull request :) 
