# VersionParser

## The problem
Excel can't sort a table based on a version number in the format #.#.#(.#.#.#...etc) properly, as this isn't a recognised format. The closest it can do is sort the values based on their string values, meaning that 1.2.11 is (incorrectly) interpreted as being less (or older) than 1.2.2.

## The solution
Some VBA that parses the version number string into a floating point number.

VersionParser contains two functions - VER2NUM and NUM2VER. Both accept up to two arguments (see Usage).

Simply parsing 1.2.3 into 1.23 wouldn't work with any revision numbers greater than 10 (as 1.2.11 -> 1.31 rather than 1.211). Additionally, even if it did parse into 1.211, it would be impossible to tell whether the original version was 1.2.11 or 1.21.1. This issue is further complicated when you consider versions involving subversions beyond the typical major.minor.patch format.

The original solution was to assign a certain number of decimal places to each subversion, with 'precision' denoting how many decimal places were assigned to each subversion - so for a precision of 3, the version xxx.yyy.zzz would be assigned to the number xxx.yyyzzz. For example 20.1.23.4 would be assigned as 20.001023004.

Mathematically speaking, each subversion would undergo the following process:
```
i = versioninteger

x = subversion

p = position of subversion (where 3 in 1.6.4.3 would be in position 4)

n = revision precision
```
<img src="https://render.githubusercontent.com/render/math?math=i=i%2B(x(10^{(-p)%2B1-(n(p-1))}))&mode=display" title="i=i+(x*(10^((p*-1)+1-(n*(p-1)))))" />

However, once I started to code a function to reverse this process, NUM2VER, I quickly ran into issues with floating-point arithmetic. To get around this issue, I simply modified the program so that instead of using powers of 10 to encode the version into a number, it uses powers of 2 instead. 

This compromises on the readability, as the parsed floating-point number is now unrecognisable as the original version string, but crucially it is still functionally identical.

## Usage

VER2NUM(version,[precision]) where version is a string and precision is an integer, will output a floating point number unique to the version for any given precision.

NUM2VER(version,[precision]) where version is a floating point number and precision is an integer, will output the version as a string corresponding to the floating point input.

Unless otherwise specified, precision will default to 4.

## Limitations
The relationship between the maximum values for the following variables can be found below.
```
r = The maximum revision number (i.e. the greatest value a subversion can take. NB: This is different from the maximum major revision number.).

m = The maximum major revision number (i.e. the greatest value the major revision number (4 in 4.2.3 or 1 in 1.6.2) can take).

v = The number of subversions (where 1.3.2 has 2 subversions, and 4.1.3.6.1 has 4 subversions).

n = The revision precision (the somewhat arbitrary variable mentioned above to switch between depth and breadth of subversions).

The maximum major revision (i.e. the first number in the version - x.0.0).
```
The following equation in combination with the restrictions below can be used to calculate whether VersionParser will be able to parse the version number correctly.

<img src="https://render.githubusercontent.com/render/math?math=m=2^{53-(v(n%2B1))}-1&mode=display" title="m=2^{53-(v(n+1))}-1" />

### Restrictions

<img src="https://render.githubusercontent.com/render/math?math=v\geq0&mode=display" title="v>=0" />

For v to be valid, m must be either greater than 0, or equal to -0.5 (the latter being an edge case where major revision number can = 0).

The major revision number must in all cases be less than 999,999,999,999,999 (1E+15), as excel cannot process any more than 15 digits of precision when performing calculations (in accordance with the [IEEE 754 floating-point standard](https://en.wikipedia.org/wiki/IEEE_754)).

<img src="https://render.githubusercontent.com/render/math?math=n\lt52&mode=display" title="n<52" />

n (the revision precision value) must be less than 52, else VersionParser will be unable to process any subversions.

### Default restrictions

For the default revision precision value (n) of 4, this will result in the following restrictions:

Maximum revision number = 15 (i.e. no subversions can be greater than revision 15 - e.g. 6.15.2 would be valid whilst 1.14.19 would be invalid (as 19 > 15).

Maximum major revision number | Number of Subversions | Final parseable version number
---|---|---
7|10|7.15.15.15.15.15.15.15.15.15.15
255|9|255.15.15.15.15.15.15.15.15.15
8191|8|8191.15.15.15.15.15.15.15.15
262143|7|262143.15.15.15.15.15.15.15
8388607|6|8388607.15.15.15.15.15.15
268435455|5|268435455.15.15.15.15.15
8589934591|4|8589934591.15.15.15.15
274877906943|3|274877906943.15.15.15
8796093022207|2|8796093022207.15.15
281474976710655|1|281474976710655.15

## Recommended settings
If you only want to parse versions following the standard major.minor.patch format, the following precision values will equate to the following maximum values.

If you are unsure what precision value you should choose for major.minor.patch, n=16 is the recommended value.

Revision Precision (n) | Maximum Major revision number (m) | Maximum revision number (r) | Total number of processable revisions | Final parseable version number
---|---|---|---|---
2|140737488355327|3|140737488355333|140737488355327.3.3
3|35184372088831|7|35184372088845|35184372088831.7.7
4|8796093022207|15|8796093022237|8796093022207.15.15
5|2199023255551|31|2199023255613|2199023255551.31.31
6|549755813887|63|549755814013|549755813887.63.63
7|137438953471|127|137438953725|137438953471.127.127
8|34359738367|255|34359738877|34359738367.255.255
9|8589934591|511|8589935613|8589934591.511.511
10|2147483647|1023|2147485693|2147483647.1023.1023
11|536870911|2047|536875005|536870911.2047.2047
12|134217727|4095|134225917|134217727.4095.4095
13|33554431|8191|33570813|33554431.8191.8191
14|8388607|16383|8421373|8388607.16383.16383
15|2097151|32767|2162685|2097151.32767.32767
16|524287|65535|655357|524287.65535.65535
17|131071|131071|393213|131071.131071.131071
18|32767|262143|557053|32767.262143.262143
19|8191|524287|1056765|8191.524287.524287
20|2047|1048575|2099197|2047.1048575.1048575
21|511|2097151|4194813|511.2097151.2097151
22|127|4194303|8388733|127.4194303.4194303
23|31|8388607|16777245|31.8388607.8388607
24|7|16777215|33554437|7.16777215.16777215
25|1|33554431|67108863|1.33554431.33554431

## Contributions
Feel free to improve this project in any way by simply submitting a Pull request :) 
