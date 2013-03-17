SuunEndo
========

This script translates Movescount Excel exports into a .tcx file that can be
imported to Endomondo.

Tech
----

Reads using `xlrd` from the `.xlsx` and writes to `.tcx` with `ETree`.

Caveats
-------

* This does not merge tracks to heart rate data
  * That means tracks in Movescount
  * That also means tracks in Endomondo
* The timezone is hardcoded to EET
