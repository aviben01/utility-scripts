# GIS Convert

Use macros this VBA module to convert long/lat coordinates into ITM East(X) / North (Y)

## Example:

**Formula:**
| lat |	long |	X |	Y |
|---|---|---|---|
| 32.7	| 35.3	| =ToItmXEast(B2,A2)	| =ToItmYNorth(B2,A2) |


**Result:**
| lat |	long |	X |	Y |
|---|---|---|---|
| 32.7	| 35.3	| 228411	| 733942 |

**Formula:**
| lat |	long |	X |	Y |
|---|---|---|---|
| 228411	| 733942 | =ItmToLongitude(B2,A2)	| =ItmToLatitude(B2,A2) |
