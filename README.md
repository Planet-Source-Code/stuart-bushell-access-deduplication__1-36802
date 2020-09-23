<div align="center">

## Access deduplication


</div>

### Description

We had a client who 'accidently' copied one databse table into itself causing duplicate records. Unfortunately Access only allows 10 columns in its De Duplicate wizard and we had twenty! This code allows you to pick the database you want deduplicating, lists the tables and the fields within the table so that you can chose the 'key' fields to deduplicate on. Then it creates a table within the database which contains all of the deleted duplicates from the table you chose.

It assumes that you don't mind which instance of a duplicate record you want to keep and it uses Jet 4.0.

Also if you run the appliaction twice on the same table you will loose all of the deleted records from the first session (it dosent create a second instance of the deleted table).

It also dosen't contain Error trapping as I use Aivostso's VB Watch addin that creates error trapping for you. However, leaving that in would add superflous code so I have omitted it here.

This code was wriiten on the quick but it has ben used in anger with correct results. Any constructive comments will be appriciated.
 
### More Info
 
The code is menu driven. Search for a database using the drive, file and directory boxes. Click on the databse and a list of tables appaer in a combo box. Click the table name and the fields are listed in a listview from which you can click the check box next to to select your fields. Then click DeDup!

A table will be created in the database you are deduping on. It will be called the same as the original table but prefixed with Deleted.


<span>             |<span>
---                |---
**Submitted On**   |2002-07-11 12:51:14
**By**             |[Stuart Bushell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stuart-bushell.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Access\_ded1048777112002\.zip](https://github.com/Planet-Source-Code/stuart-bushell-access-deduplication__1-36802/archive/master.zip)








