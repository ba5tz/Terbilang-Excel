# Terbilang
Fungsi Terbilang Excel digunakan untuk mengubah Nilai Angka (Number) menjadi Text terbilang dalam Bahasa Indonesia.

![alt text](images.PNG?raw=true "SC")

- [Versi Formula](#terbilang-versi-formula)
- [versi UDF](#terbilang-udf)
- [Versi Addin](#versi-add-in)
- [Author](#author)

## Terbilang Versi Formula 
Ini adalah versi paling mudah, caranya cukup copy rumus ini ke B1, silahkan sesuaikan dengan Regional Setting Excel yang digunakan. Atau silahkan sesuaikan Range sesuai dengan lokasi nilai yang ingin diubah menjadi terbilang.
### (Regional Setting Indonesia)
```js
IF(A1=0;"nol";IF(A1<0;"minus ";"")& SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE( IF(--MID(TEXT(ABS(A1);"000000000000000");1;3)=0;"";MID(TEXT(ABS(A1);"000000000000000");1;1)&" ratus "&MID(TEXT(ABS(A1);"000000000000000");2;1)&" puluh "&MID(TEXT(ABS(A1);"000000000000000");3;1)&" trilyun ")& IF(--MID(TEXT(ABS(A1);"000000000000000");4;3)=0;"";MID(TEXT(ABS(A1);"000000000000000");4;1)&" ratus "&MID(TEXT(ABS(A1);"000000000000000");5;1)&" puluh "&MID(TEXT(ABS(A1);"000000000000000");6;1)&" milyar ")& IF(--MID(TEXT(ABS(A1);"000000000000000");7;3)=0;"";MID(TEXT(ABS(A1);"000000000000000");7;1)&" ratus "&MID(TEXT(ABS(A1);"000000000000000");8;1)&" puluh "&MID(TEXT(ABS(A1);"000000000000000");9;1)&" juta ")& IF(--MID(TEXT(ABS(A1);"000000000000000");10;3)=0;"";IF(--MID(TEXT(ABS(A1);"000000000000000");10;3)=1;"*";MID(TEXT(ABS(A1);"000000000000000");10;1)&" ratus "&MID(TEXT(ABS(A1);"000000000000000");11;1)&" puluh ")&MID(TEXT(ABS(A1);"000000000000000");12;1)&" ribu ")& IF(--MID(TEXT(ABS(A1);"000000000000000");13;3)=0;"";MID(TEXT(ABS(A1);"000000000000000");13;1)&" ratus "&MID(TEXT(ABS(A1);"000000000000000");14;1)&" puluh "&MID(TEXT(ABS(A1);"000000000000000");15;1));1;"satu");2;"dua");3;"tiga");4;"empat");5;"lima");6;"enam");7;"tujuh");8;"delapan");9;"sembilan");"0 ratus";"");"0 puluh";"");"satu puluh 0";"sepuluh");"satu puluh satu";"sebelas");"satu puluh dua";"duabelas");"satu puluh tiga";"tigabelas");"satu puluh empat";"empatbelas");"satu puluh lima";"limabelas");"satu puluh enam";"enambelas");"satu puluh tujuh";"tujuhbelas");"satu puluh delapan";"delapanbelas");"satu puluh sembilan";"sembilanbelas");"satu ratus";"seratus");"*satu ribu";"seribu");0;""));" ";" "))
```

### (Regional Setting US)
```js
=IF(A1=0,"nol",IF(A1<0,"minus ","")& SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE( IF(--MID(TEXT(ABS(A1),"000000000000000"),1,3)=0,"",MID(TEXT(ABS(A1),"000000000000000"),1,1)&" ratus "&MID(TEXT(ABS(A1),"000000000000000"),2,1)&" puluh "&MID(TEXT(ABS(A1),"000000000000000"),3,1)&" trilyun ")& IF(--MID(TEXT(ABS(A1),"000000000000000"),4,3)=0,"",MID(TEXT(ABS(A1),"000000000000000"),4,1)&" ratus "&MID(TEXT(ABS(A1),"000000000000000"),5,1)&" puluh "&MID(TEXT(ABS(A1),"000000000000000"),6,1)&" milyar ")& IF(--MID(TEXT(ABS(A1),"000000000000000"),7,3)=0,"",MID(TEXT(ABS(A1),"000000000000000"),7,1)&" ratus "&MID(TEXT(ABS(A1),"000000000000000"),8,1)&" puluh "&MID(TEXT(ABS(A1),"000000000000000"),9,1)&" juta ")& IF(--MID(TEXT(ABS(A1),"000000000000000"),10,3)=0,"",IF(--MID(TEXT(ABS(A1),"000000000000000"),10,3)=1,"*",MID(TEXT(ABS(A1),"000000000000000"),10,1)&" ratus "&MID(TEXT(ABS(A1),"000000000000000"),11,1)&" puluh ")&MID(TEXT(ABS(A1),"000000000000000"),12,1)&" ribu ")& IF(--MID(TEXT(ABS(A1),"000000000000000"),13,3)=0,"",MID(TEXT(ABS(A1),"000000000000000"),13,1)&" ratus "&MID(TEXT(ABS(A1),"000000000000000"),14,1)&" puluh "&MID(TEXT(ABS(A1),"000000000000000"),15,1)),1,"satu"),2,"dua"),3,"tiga"),4,"empat"),5,"lima"),6,"enam"),7,"tujuh"),8,"delapan"),9,"sembilan"),"0 ratus",""),"0 puluh",""),"satu puluh 0","sepuluh"),"satu puluh satu","sebelas"),"satu puluh dua","duabelas"),"satu puluh tiga","tigabelas"),"satu puluh empat","empatbelas"),"satu puluh lima","limabelas"),"satu puluh enam","enambelas"),"satu puluh tujuh","tujuhbelas"),"satu puluh delapan","delapanbelas"),"satu puluh sembilan","sembilanbelas"),"satu ratus","seratus"),"*satu ribu","seribu"),0,""))," "," "))
```

## Terbilang UDF
Cara menggunakan versi UDF, silahka copy Script ke Module (Visual basic Editor - Insert module). gunakan rumus terbilang seperti rumus Excel pada umumnya dengan mengetikan =Terbilang(A1) jika Nilai berada di A1.
```vbs
Function Terbilang(n As Long) As String 'max 2.147.483.647
Dim satuan As Variant, Minus As Boolean
On Error GoTo terbilang_error
satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
If n < 0 Then
    Minus = True
    n = n * -1
End If
Select Case n
    Case 0 To 11
        Terbilang = " " + satuan(Fix(n))
    Case 12 To 19
        Terbilang = Terbilang(n Mod 10) + " Belas"
    Case 20 To 99
        Terbilang = Terbilang(Fix(n / 10)) + " Puluh" + Terbilang(n Mod 10)
    Case 100 To 199
        Terbilang = " Seratus" + Terbilang(n - 100)
    Case 200 To 999
        Terbilang = Terbilang(Fix(n / 100)) + " Ratus" + Terbilang(n Mod 100)
    Case 1000 To 1999
        Terbilang = " Seribu" + Terbilang(n - 1000)
    Case 2000 To 999999
        Terbilang = Terbilang(Fix(n / 1000)) + " Ribu" + Terbilang(n Mod 1000)
    Case 1000000 To 999999999
        Terbilang = Terbilang(Fix(n / 1000000)) + " Juta" + Terbilang(n Mod 1000000)
    Case Else
        Terbilang = Terbilang(Fix(n / 1000000000)) + " Milyar" + Terbilang(n Mod 1000000000)
End Select
If Minus = True Then
    Terbilang = "Minus" + Terbilang
End If
Exit Function
terbilang_error:
MsgBox Err.Description, vbCritical, "Terbilang Error"
End Function
```

## Versi Add-in
Untuk versi Addin, silahkan download dan Install Addin pada Microsoft Excel yang digunakan. Terbilang bisa langsung digunakan dengan menuliskan Rumus =Terbilang(A1) jika Nilai berada di A1
- Download URL: [Terbilang Excel Add-in](https://www.excelnoob.com/formula-ms-excel-terbaru-dalam-addin-udf/)


## Author
[![Author](https://img.shields.io/badge/author-Andi%20B.%20Setiadi-lightgrey.svg?colorB=1D63DC&style=flat-square)]()

Thanks
    - More Info [setiadi.my.id](https://setiadi.my.id)
    - Formula terbilang By [KelasExcel.id](https://kelasexcel.id) 
    - VBA Script Terbilang by [Excelnoob.com](https://excelnoob.com)
