# Excel Hash

![img](./Images/excel-hash.png)

This repository contains an exercise I have created to study the concept of [Hashing Algorithms](https://www.sciencedirect.com/topics/computer-science/hashing-algorithm#:~:text=Cryptographic%20hashing%20algorithms%2C%20also%20known,product%2C%20called%20the%20hash%20value.) in cyber security. 

The exercise consists of processing a given hypotetical Excel table which contains information that the author doesn't want anyone else to modify, but view without having to use the traditional solution of locking the spreadsheet from editing by using a password (https://bit.ly/3Wi3Wyd). In this exercise, the Author you will be able to share the spreadsheet with the Reviewers and check later whether it was modified by using macro VBScripts programmed with a simple hash algorithm.

## The source table

The source Excel table is available in the spreadsheet ![Excel_Hash](Excel_Hash.xlsm) and consists of a hypothetic list of people's information[^1], as shown below:

![table](/Images/table-sample.png)

## The hashing logic

**The hashing logic in this exercise is as follows:**

**Step 1.** The Author creates the spreadsheet containing the table above.

**Step 2.** The Author runs a **hash creation algorithm** which creates a hash code from every line in the table.

**Step 3.** The Author saves the spreadsheet and sends it for the Reviewers.

**Step 4.** The Reviewers open the spreadsheet and review the information. In this hypothetical scenario, the Reviewers can make notes about their reviews in a separate document or even in the body of the email, **but never** in the spreadsheet, nor changing any of its information to meet their requirements (I know this looks unreal, but creating a realistic scenario is not my target here, but the hashing exercise itself).

**Step 5.** Each Reviewer sends the spreadsheet back to the Author containing its notes (either in the email or a separate document).

**Step 6.** The Author runs a **hash review algorithm** which will verify whether the data in the table has been changed and inform the Author.

**Let's see how it looks in practice:**

**Step 2:** By pressing ALT+F8 the Author can access the Excel's interface for macro execution. Then, he must select the **GenerateHash** macro and click on *Execute*.

![image](/Images/Generate-Hash-2.png)

The macro script will then executes the hashing algorithm and inform whether the process was or not performed sucessfully:

![image](/Images/hash-completed-2.png)

The codes are always generated in the second column after the final column of data in the spreadsheet (in the example, column H):

![image](/Images/hash-codes.png)

The codes are formatted in white color so that they don't pollute the spreadsheet and also become invisible for the Reviewers.
Notice that the hash codes have the same size of 12 characters (predefined by me), regardless of the size of the texts used as input. This makes the hashing algorithms an excellent ally when hiding information. So both "Hello" and "Hi, my name is John" will generate a hash code with the same size predefined.


[^1]: You can see that I am using only confidential information, such as address, credit card number, password etc., which looks weird, but this is because I have the intention to use this table for another exercise about cryptography.

**Step 6:** After receiving the spreadsheet from a Reviewer, the Author must again press ALT+F8 to access the Excel's interface for macro execution. This time, he must select the **VerifyHash** macro and click on *Execute*. This macro will run the same hashing algorithm as the previous **GenerateHash** macro, but this time making an additional comparison between each hash code generated against the previous code stored in the column H. Since the first difference is found, the process will be interrupted and the user informed.

In case of the spreadsheet hasn't been changed, the macro will show a sucessfully message:

![image](/Images/verify-hash-sucess.png)

Now let's see how it works when verifying a spreadsheet which has been changed. For this case, I have done small changes in the lines 6, 8, and 10, in the information highlighted in red in the image below:

![image](/Images/changes-in-data.png)

Some of these changes were really small, like replacing a single number! But as shown below, the algorithm sucessfully identifies those changes proving the efficiency of hashing to this type of investigation:

![image](/Images/verify-hash-alert.png)

## The VBScripts

Now let's take a look into the VBScripts that implements the hashing algorithms. I've put many comments in the code, so I hope they are self explanatory.

### Sub GenerateHash()
This is the script that generates the hash codes by performing the main logic and using the SimpleHash() function as the core for the hashing.

```VBScript
Sub GenerateHash()
'Generates the hash values from the spreadsheet original values and
'stores the hashes at two columns from the last column of data, formatted
'in white color so that they don't appear.
    
    On Error GoTo ErrorHandler ' Set up error handling
    Dim wsOriginal As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long, c As Long
    Dim cellValue As String
    Dim concatValue As String
    Dim hashedValue As String
    
    ' Sets the source spreadsheet (adjust as necessary)
    Set wsOriginal = ThisWorkbook.Sheets("Records")
    
    ' Finds the last line and column of the source table
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    
    ' Applies the hash algorithm to the data and records the results at two positions after the last column
    For r = 2 To lastRow
        concatValue = ""
        cellValue = ""
        For c = 1 To lastCol
           cellValue = wsOriginal.Cells(r, c).Value
           hashedValue = SimpleHash(cellValue)
           concatValue = concatValue & hashedValue
        Next c
        wsOriginal.Cells(r, lastCol + 2).Value = "'" & concatValue ' Adds an apostrophe to force the value to be text
        wsOriginal.Cells(r, lastCol + 2).Font.Color = RGB(255, 255, 255) ' Defines the font color as white
    Next r
    ' Display a success message
    MsgBox "Hash generation completed successfully!", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    ' Handle any errors that occur
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub
```
### Function SimpleHash()
This function contains the core logic of the hashing process. For simplification, I decided to use a very simple logic to calculate the hash codes based on the ASCII values of each character. You can learn more about real hashing algorithms by reading the [Secure Hash Algorithms](https://en.wikipedia.org/wiki/Secure_Hash_Algorithms).

```VBScript
Function SimpleHash(inputString As String) As String
    Dim i As Long
    Dim charValue As Long
    Dim total As Long
    Dim hash As String
    
    total = 0
    
    ' Calculates the total value of the characters of a string
    For i = 1 To Len(inputString)
        charValue = Asc(Mid(inputString, i, 1)) ' Converts the extracted character to its ASCII value.
        total = total + charValue
    Next i
    
    ' Converts the total into a hash string
    hash = ""
    Do While total > 0
        ' Computes the remainder of total divided by 94. This ensures the resulting value is within the range of 0 to 93.
        ' Converts the remainder to a character. Adding 33 shifts the value to the printable
        ' ASCII character range (from '!' (33) to '~' (126)).
        hash = hash & Chr((total Mod 94) + 33)
        ' Uses integer division to reduce total, effectively moving to the next digit.
        total = total \ 94
    Loop
    
    SimpleHash = hash
End Function
```
### Sub VerifyHash()
```VBScript
Sub VerifyHash()
' Verifies the hash values in the original spreadsheet original values and
' stores the hashes at two columns from the last column of data, formatted
' in white color so that they don't appear.
  
    Dim wsOriginal As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long, c As Long
    Dim cellValue As String
    Dim concatValue As String
    Dim recalcHash As String
    Dim storedHash As String
    Dim isCorrect As Boolean
    Dim mismatchedRows As String
    
    ' Sets the source spreadsheet (adjust as necessary)
    Set wsOriginal = ThisWorkbook.Sheets("Records")
    
    ' Finds the last line and column of the source table
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    
    ' Considers that the original hash values are stored at two columns after the last data column
    Dim hashCol As Long
    hashCol = lastCol + 2
    
    ' Verifies the hashs integrity
    isCorrect = True
    mismatchedRows = ""
    For r = 2 To lastRow ' Considering the first row always has a header
        concatValue = ""
        For c = 1 To (lastCol)
           cellValue = wsOriginal.Cells(r, c).Value
           recalcHash = SimpleHash(cellValue)
           concatValue = concatValue & recalcHash
        Next c
        storedHash = wsOriginal.Cells(r, hashCol).Value
        If concatValue <> storedHash Then
            isCorrect = False
            mismatchedRows = mismatchedRows & "Row " & r & "; "
        End If
    Next r
    
    ' Shows a message with the result of the verification
    If isCorrect Then
        MsgBox "The information has not been changed.", vbInformation
    Else
      ' Remove the trailing semicolon and space from the mismatchedCells string
      If Len(mismatchedRows) > 0 Then
        mismatchedRows = Left(mismatchedRows, Len(mismatchedRows) - 2)
      End If
      MsgBox "The information has been changed in the following rows: " & mismatchedRows, vbExclamation
    End If
End Sub
