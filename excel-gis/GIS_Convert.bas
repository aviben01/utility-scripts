Option Explicit


Const PI As Double = 3.14159265358979

Type DATUM
    A As Double     'a  Equatorial earth radius
    b As Double     'b  Polar earth radius
    f As Double     'f= (a-b)/a  Flatenning
    esq As Double   'esq = 1-(b*b)/(a*a)  Eccentricity Squared
    E As Double     'sqrt(esq)  Eccentricity
    ' deltas to WGS84
    dX As Double
    dY As Double
    dZ As Double
End Type

Type GRID
    lon0 As Double
    lat0 As Double
    k0 As Double
    false_e As Double
    false_n As Double
End Type

' WGS84 data
Function GetWgs84Data() As DATUM
    Dim Data As DATUM
    Data.A = 6378137#
    Data.b = 6356752.3142
    Data.f = 1 / 298.257223563
    Data.esq = 6.69438000426081E-03
    Data.E = 8.18191909289062E-02
    ' deltas to WGS84
    Data.dX = 0#
    Data.dY = 0#
    Data.dZ = 0#
    GetWgs84Data = Data
End Function

' GRS80 data
Function GetGrs80Data() As DATUM
    Dim Data As DATUM
    Data.A = 6378137#
    Data.b = 6356752.3141
    Data.f = 1 / 298.257222101
    Data.esq = 6.69438002290272E-03
    Data.E = 8.18191910428276E-02
    ' deltas to WGS84
    Data.dX = -48#
    Data.dY = 55#
    Data.dZ = 52#
    GetGrs80Data = Data
End Function

' Clark 1880 Modified data
Function GetClark1880ModifiedData() As DATUM
    Data As DATUM
    Data.A = 6378300.789
    Data.b = 6356566.4116309
    Data.f = 1 / 293.466
    Data.esq = 6.80348813911232E-03
    Data.E = 8.24832597507659E-02
    ' deltas to WGS84
    Data.dX = -235#
    Data.dY = -85#
    Data.dZ = 264#
    GetClark1880ModifiedData = Data
End Function
    
' ICS data
Function GetIcsDataGrid() As GRID
    Dim Grd As GRID
    Grd.lon0 = 0.6145667421719          ' lon0 = central meridian in radians of 35.12'43.490"
    Grd.lat0 = 0.553864476827628        ' lat0 = central latitude in radians of 31.44'02.749"
    Grd.k0 = 1#                         ' k0 = scale factor
    Grd.false_e = 170251.555            ' false_easting
    Grd.false_n = 2385259#              ' false_northing
    GetIcsDataGrid = Grd
End Function

' ITM data
Function GetItmDataGrid() As GRID
    Dim Grd As GRID
    Grd.lon0 = 0.614434732254689        '  lon0 = central meridian in radians 35.12'16.261"
    Grd.lat0 = 0.553869654637742        '  lat0 = central latitude in radians 31.44'03.817"
    Grd.k0 = 1.0000067                  '  k0 = scale factor
    Grd.false_e = 219529.584            '  false_easting
    Grd.false_n = 2885516.9488          '  false_northing = 3512424.3388-626907.390
                                        ' MAPI says the false northing is 626907.390, and in another place
                                        ' that the meridional arc at the central latitude is 3512424.3388
    GetItmDataGrid = Grd
End Function

Function Pow(ByVal X As Double, ByVal N As Double) As Double
    Pow = X ^ N
End Function

Sub itm2wgs84(N As Double, E As Double, ByRef lat As Double, ByRef lon As Double)
    Dim lat80 As Double, lon80 As Double
    Dim lat84 As Double, lon84 As Double
    
    Call Grid2LatLon(N, E, lat80, lon80, GetItmDataGrid(), GetGrs80Data())
    Call Molodensky(lat80, lon80, lat84, lon84, GetGrs80Data(), GetWgs84Data())
    lat = lat84 * 180# / PI
    lon = lon84 * 180# / PI
End Sub

Sub wgs842itm(lat As Double, lon As Double, ByRef N As Double, ByRef E As Double)
    Dim latr As Double: latr = lat * PI / 180#
    Dim lonr As Double: lonr = lon * PI / 180#
    Dim lat80 As Double, lon80 As Double
    Call Molodensky(latr, lonr, lat80, lon80, GetWgs84Data(), GetGrs80Data())
    Call LatLon2Grid(lat80, lon80, N, E, GetGrs80Data(), GetItmDataGrid())
End Sub

Sub ics2wgs84(N As Double, E As Double, ByRef lat As Double, ByRef lon As Double)
    Dim lat80, lon80 As Double
    Call Grid2LatLon(N, E, lat80, lon80, GetIcsDataGrid(), GetClark1880ModifiedData())
    Dim lat84, lon84 As Double
    Call Molodensky(lat80, lon80, lat84, lon84, GetClark1880ModifiedData(), GetWgs84Data())
    lat = lat84 * 180# / PI
    lon = lon84 * 180# / PI
End Sub

Sub wgs842ics(lat As Double, lon As Double, ByRef N As Double, ByRef E As Double)
    Dim latr As Double: latr = lat * PI / 180#
    Dim lonr As Double: lonr = lon * PI / 180#
    Dim lat80, lon80 As Double
    Call Molodensky(latr, lonr, lat80, lon80, GetWgs84Data(), GetClark1880ModifiedData())
    Call LatLon2Grid(lat80, lon80, N, E, GetClark1880ModifiedData(), GetIcsDataGrid())
End Sub

Sub Grid2LatLon(ByVal N As Double, ByVal E As Double, ByRef lat As Double, ByRef lon As Double, ByRef from_grid As GRID, ByRef to_data As DATUM)
    Dim Y As Double:    Y = N + from_grid.false_n
    Dim X As Double:    X = E - from_grid.false_e
    Dim M As Double:    M = Y / from_grid.k0
    Dim A As Double:    A = to_data.A
    Dim b As Double:    b = to_data.b
    E = to_data.E
    Dim esq As Double:  esq = to_data.esq
    Dim mu As Double:   mu = M / (A * (1 - E * E / 4 - 3 * Pow(E, 4) / 64 - 5 * Pow(E, 6) / 256))
    Dim ee As Double:   ee = Math.Sqr(1 - esq)
    Dim e1 As Double:   e1 = (1 - ee) / (1 + ee)
    Dim j1 As Double:   j1 = 3 * e1 / 2 - 27 * e1 * e1 * e1 / 32
    Dim j2 As Double:   j2 = 21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32
    Dim j3 As Double:   j3 = 151 * e1 * e1 * e1 / 96
    Dim j4 As Double:   j4 = 1097 * e1 * e1 * e1 * e1 / 512
    Dim fp As Double:   fp = mu + j1 * Math.Sin(2 * mu) + j2 * Math.Sin(4 * mu) + j3 * Math.Sin(6 * mu) + j4 * Math.Sin(8 * mu)
    Dim sinfp As Double:    sinfp = Math.Sin(fp)
    Dim cosfp As Double:    cosfp = Math.Cos(fp)
    Dim tanfp As Double:    tanfp = sinfp / cosfp
    Dim eg As Double:   eg = (E * A / b)
    Dim eg2 As Double:  eg2 = eg * eg
    Dim C1 As Double:   C1 = eg2 * cosfp * cosfp
    Dim T1 As Double:   T1 = tanfp * tanfp
    Dim R1 As Double:   R1 = A * (1 - E * E) / Pow(1 - (E * sinfp) * (E * sinfp), 1.5)
    Dim N1 As Double:   N1 = A / Math.Sqr(1 - (E * sinfp) * (E * sinfp))
    Dim D As Double:    D = X / (N1 * from_grid.k0)
    Dim Q1 As Double:   Q1 = N1 * tanfp / R1
    Dim Q2 As Double:   Q2 = D * D / 2
    Dim Q3 As Double:   Q3 = (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * eg2 * eg2) * (D * D * D * D) / 24
    Dim Q4 As Double:   Q4 = (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 3 * C1 * C1 - 252 * eg2 * eg2) * (D * D * D * D * D * D) / 720
    lat = fp - Q1 * (Q2 - Q3 + Q4)
    Dim Q5 As Double:   Q5 = D
    Dim Q6 As Double:   Q6 = (1 + 2 * T1 + C1) * (D * D * D) / 6
    Dim Q7 As Double:   Q7 = (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * eg2 * eg2 + 24 * T1 * T1) * (D * D * D * D * D) / 120
    lon = from_grid.lon0 + (Q5 - Q6 + Q7) / cosfp
End Sub

Sub LatLon2Grid(ByVal lat As Double, ByVal lon As Double, ByRef N As Double, ByRef E As Double, ByRef from_data As DATUM, ByRef to_data As GRID)
    E = from_data.E
    Dim A As Double:    A = from_data.A
    Dim b As Double:    b = from_data.b
    Dim slat1 As Double:    slat1 = Math.Sin(lat)
    Dim clat1 As Double:    clat1 = Math.Cos(lat)
    Dim clat1sq As Double:  clat1sq = clat1 * clat1
    Dim tanlat1sq As Double:    tanlat1sq = slat1 * slat1 / clat1sq
    Dim e2 As Double:   e2 = Pow(E, 2)
    Dim e4 As Double:   e4 = Pow(E, 4)
    Dim e6 As Double:   e6 = Pow(E, 6)
    Dim eg As Double:   eg = (E * A / b)
    Dim eg2 As Double:  eg2 = Pow(eg, 2)
    Dim l1 As Double:   l1 = 1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256
    Dim l2 As Double:   l2 = 3 * e2 / 8 + 3 * e4 / 32 + 45 * e6 / 1024
    Dim l3 As Double:   l3 = 15 * e4 / 256 + 45 * e6 / 1024
    Dim l4 As Double:   l4 = 35 * e6 / 3072
    Dim M As Double:    M = A * (l1 * lat - l2 * Math.Sin(2 * lat) + l3 * Math.Sin(4 * lat) - l4 * Math.Sin(6 * lat))
    Dim nu As Double:   nu = A / Math.Sqr(1 - (E * slat1) * (E * slat1))
    Dim p As Double:    p = lon - to_data.lon0
    Dim k0 As Double:   k0 = to_data.k0
    Dim K1 As Double:   K1 = M * k0
    Dim K2 As Double:   K2 = k0 * nu * slat1 * clat1 / 2
    Dim K3 As Double:   K3 = (k0 * nu * slat1 * clat1 * clat1sq / 24) * (5 - tanlat1sq + 9 * eg2 * clat1sq + 4 * eg2 * eg2 * clat1sq * clat1sq)
    Dim Y As Double:    Y = K1 + K2 * p * p + K3 * p * p * p * p - to_data.false_n
    Dim K4 As Double:   K4 = k0 * nu * clat1
    Dim K5 As Double:   K5 = (k0 * nu * clat1 * clat1sq / 6) * (1 - tanlat1sq + eg2 * clat1 * clat1)
    Dim X As Double:    X = K4 * p + K5 * p * p * p + to_data.false_e
    E = CLng((X + 0.5))
    N = CLng((Y + 0.5))
End Sub

Sub Molodensky(ByVal ilat As Double, ByVal ilon As Double, ByRef olat As Double, ByRef olon As Double, ByRef from_data As DATUM, ByRef to_data As DATUM)
    Dim dX As Double:   dX = from_data.dX - to_data.dX
    Dim dY As Double:   dY = from_data.dY - to_data.dY
    Dim dZ As Double:   dZ = from_data.dZ - to_data.dZ
    Dim slat As Double: slat = Math.Sin(ilat)
    Dim clat As Double: clat = Math.Cos(ilat)
    Dim slon As Double: slon = Math.Sin(ilon)
    Dim clon As Double: clon = Math.Cos(ilon)
    Dim ssqlat As Double:   ssqlat = slat * slat
    Dim from_f As Double:   from_f = from_data.f
    Dim df As Double:   df = to_data.f - from_f
    Dim from_a As Double:   from_a = from_data.A
    Dim da As Double:   da = to_data.A - from_a
    Dim from_esq As Double: from_esq = from_data.esq
    Dim adb As Double:  adb = 1# / (1# - from_f)
    Dim rn As Double:   rn = from_a / Math.Sqr(1 - from_esq * ssqlat)
    Dim rm As Double:   rm = from_a * (1 - from_esq) / Pow((1 - from_esq * ssqlat), 1.5)
    Dim from_h As Double:   from_h = 0#
    Dim dlat As Double: dlat = (-dX * slat * clon - dY * slat * slon + dZ * clat + da * rn * from_esq * slat * clat / from_a + df * (rm * adb + rn / adb) * slat * clat) / (rm + from_h)
    olat = ilat + dlat
    Dim dlon As Double: dlon = (-dX * slon + dY * clon) / ((rn + from_h) * clat)
    olon = ilon + dlon
End Sub


Public Function ToItmYNorth(lon As Double, lat As Double) As Double
    Dim N As Double, E As Double
    Call wgs842itm(lat, lon, N, E)
    ToItmYNorth = Math.Round(N, 2)
End Function

Public Function ToItmXEast(lon As Double, lat As Double) As Double
    Dim N As Double, E As Double
    Call wgs842itm(lat, lon, N, E)
    ToItmXEast = Math.Round(E, 2)
End Function

Public Function ItmToLatitude(N As Double, E As Double) As Double
    Dim lon As Double, lat As Double
    Call itm2wgs84(N, E, lat, lon)
    ItmToLatitude = Math.Round(lat, 6)
End Function

Public Function ItmToLongitude(N As Double, E As Double) As Double
    Dim lon As Double, lat As Double
    Call itm2wgs84(N, E, lat, lon)
    ItmToLongitude = Math.Round(lon, 6)
End Function

Sub conversion()
    Dim X As Double, Y As Double
    Dim N As Double, E As Double
    
    X = 35.234383
    Y = 31.776748
    X = 35
    Y = 32
    
    N = ToItmYNorth(X, Y)
    E = ToItmXEast(X, Y)
    
    Debug.Print "ToItm", X, Y, E, N
    
    X = 0: Y = 0
    X = ItmToLongitude(N, E)
    Y = ItmToLatitude(N, E)
    Debug.Print "FromItm", X, Y, E, N

End Sub

