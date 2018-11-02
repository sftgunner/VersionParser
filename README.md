# VersionParser

## The problem
Excel can't sort a table based on a version number in the format #.#.#(.#.#.#...etc) properly, as this isn't a recognised format. The closest it can do is sort the values based on their string values, meaning that 1.2.11 is (incorrectly) interpreted as being less (or an older version) than 1.2.2.

## The solution
Write a bit of VBA that parses the version number string into a floating point number.

In fact, it's (slightly) cleverer than it sounds, as simply parsing 1.2.3 into 1.23 wouldn't work with any revision numbers greater than 10 (as 1.2.11 -> 1.31 rather than 1.211). Additionally, even if it did parse into 1.211, it would be impossible to tell whether the original version was 1.2.11 or 1.2.1.1

As hinted by the previous sentence, VersionParser (or VER2NUM as the function is labelled) can also cope with versioning levels beyond the typical Major.Minor.Patch format.

In fact, I've even added a NUM2VER to reverse the process! This spurred on the change from powers of 10 to powers of 2 that are used, to avoid floating point errors.

I've now (sortof) fixed one of the limitations - precision is now editable by the user, allowing the flexibility to focus on either depth (e.g. 1.2.3.4.5.6.7.8.9) or breadth (e.g. 45.102048.17842). Higher "precision"* equates to higher breadth - so a "precision" of 5 allows for 2^5 revisions per subversion, but due to the space required for this higher precision, you will only be able to use a depth of log_(2^5)(1.099e+12) = 8 subversions before the resolution is affected, and you can't convert the number back again.

*If you can come up with a better name for this than "precision", please do open a pull request/issue! I realise the terminology is somewhat confusing - I will update this README if I can work out the correct terminology to be used.

## Limitations
Aside from the bad formatting/layout/practices that I've presumably used in what is my first bit of VBA, due to limitations in excel, the maximum total revisions = ~1.099 trillion subversions, which can be configured in a variety of formats. However, please remember that you can have an unlimited number of releases/top-level versions. This is unaffected by the precision setting.

Also, yes, I realise I really ought to comment more.

## Contributions
Feel free to improve this project in any way by simply submitting a Pull request :) 
