# DND weather generator
![video](https://raw.githubusercontent.com/foreveryone1/dndweathertable/main/weathergen.webp)

Get the generator [here](https://github.com/foreveryone1/dndweathertable/raw/main/Weathergen.xlsm).

This table uses a [Lehmer pseudo-random number generator](https://en.wikipedia.org/wiki/Lehmer_random_number_generator) to generate numbers from a randomly generated or manually inputted seed.

## Using the generator

![image](https://user-images.githubusercontent.com/27033050/116296647-bfd2d680-a79a-11eb-8e37-1291156d9e78.png)

Select the climate conditions from the drop-down bar as preferred. Input a seed or press New Seed to generate a new seed.

The initialise button will choose assign a random day as well as random starting conditions. If you prefer, you can instead manually choose a starting day and starting conditions.

## Extending the generator

You can add your own climates with unique weather conditions by modifying the coloured tables in the Temp, Weather, and Wind sheets. 

Right click anywhere within a table in the aforementioned sheets and select insert table columns to add a climate type.

![image](https://user-images.githubusercontent.com/27033050/116298255-8602cf80-a79c-11eb-97c2-ce3afcb940a9.png)

When filling in a climate type, keep in mind that the weather occurrences in the middle are more likely than the events the beginning or end of the column.

On the days sheet you can also rename the days as you prefer, or even add additional days by right clicking.


## Macros Used
As this weathergen utilises Macros for quality of life I felt it best to disclose what macro functions are used to assuage any concerns.

**Generate seed**
```VBA
Sub Generate_seed()
'
' Generate_seed Macro
'

'
    Calculate
    Sheets("General sheet").Range("K17").Copy
    Sheets("General sheet").Range("M16").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
```

**Initialise**
```VBA
Sub Initialise()
'
' Initialise Macro
'

'
    Sheets("Days").Range("B1").Copy
    Sheets("General sheet").Range("B3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Temp").Range("A2").Copy
    Sheets("General sheet").Range("E3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Precip").Range("A2").Copy
    Sheets("General sheet").Range("F3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Wind").Range("A2").Copy
    Sheets("General sheet").Range("G3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
```
