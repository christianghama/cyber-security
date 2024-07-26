# Excel Hash
This repository contains an exercise I have created to study the concept of [Hashing Algorithms](https://www.sciencedirect.com/topics/computer-science/hashing-algorithm#:~:text=Cryptographic%20hashing%20algorithms%2C%20also%20known,product%2C%20called%20the%20hash%20value.) in cyber security. 

The exercise consists of processing a given hypotetical Excel table which contains information that the author doesn't want anyone else to modify, but view without having to use the traditional solution of locking the spreadsheet from editing by using a password (https://bit.ly/3Wi3Wyd), in this exercise the author you will be able to share the spreadsheet with the reviewers and check later whether it was modified by using macro VBScripts programmed with a simple hash algorithm.

## The source table

The source Excel table is available in the spreadsheet ![Excel_Hash](Excel_Hash.xlsm) and consists of a hypothetic list of people's information[^1], as shown below:

![table](/Images/table-sample.png)

## The hashing logic

The hashing logic in this exercise is as follows:

**Step 1.** The Author creates the spreadsheet containing the table above.

**Step 2.** The Author runs a **hash creation algorithm** which creates a hash code from every line in the table.

**Step 3.** The Author saves the spreadsheet and sends it for the Reviewers.

**Step 4.** The Reviewers open the spreadsheet and review the information. In this hypothetical scenario, the Reviewers can make notes about their reviews in a separate document or even in the body of the email, **but never** in the spreadsheet, nor changing any of its information to meet their requirements (I know this looks unreal, but creating a realistic scenario is not my target here, but the hashing exercise itself).

**Step 5.** Each Reviewer sends the spreadsheet back to the Author containing its notes (either in the email or a separate document).

**Step 6.** The Author runs a **hash review algorithm** which will verify whether the data in the table has been changed and inform the Author.

Let's see how it looks in practice:

**Step 2:** By pressing ALT+F8 the Author can access the Excel's interface for macro execution. Then, he must select the **GenerateHash** macro and click on *Execute*.

![macro](/Images/Generate-Hash.png)

The macro script will then executes the hashing algorithm and inform whether the process was or not performed sucessfully:

![macro](/Images/hash-completed.png)


[^1]: You can see that I am using only confidential information, such as address, credit card number, password etc., which looks weird, but this is because I have the intention to use this table for another exercise about cryptography.


