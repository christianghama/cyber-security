# Excel Hash
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

![image](/Images/changes-in-data.png)



