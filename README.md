# xl_unCode
return alpha-3 country code for country names submitted to Wolfram|Alpha API

## Wolfram|Alpha API
xl_unCode uses the Wolfram|Alpha API to return "UN code" alpha-3 country code from permutations of country names.  It requires an appid to function.  Sign up at http://products.wolframalpha.com/api/ for an account and add an app to your account to obtain an appid.

## Directions
*For Excel 2013 on Windows 8.1*

1. Enable Developer Tab.  Go to File > Options > Customize Ribbon: Main Tabs: check the Developer box
2. Open the Visual Basic Editor
3. Select the open workbook and open an empty module (Module1)
4. Copy xl_unCode.bas into module
5. insert your appid in the "" on line 21
  
  ```
  'Wolfram Alpha appid
  Dim appid As String
  appid = "xxxxxx-xxxxxxxxxx"
  ```
  
6. enable references in the Visual Basic editor.  Tools > References...
  
  a. Microsoft HTML Object Library
  
  b. Microsoft Internet Controls
  
## Usage

Make Column A the country name field in your data.  insert an empty column for column B.  This macro will fill column be regardless of what is in it.  Modify the code to change the input and output columns.

The first row is assumed to be a header.  To return a code result for the first line, change `For i = 2 To Last` to `For i = 1 To Last`

