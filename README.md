# JExcel
Build Excel via Javascript

---

```js

var jExc = new JExcel();

//Init a Sheet with a name.
var sheet = jExc.SetSheet("Sheet1");

//the row and col start at 0 . 
//Set the D4 position to a number 1000.
sheet.Set(3, 3, 1000);

//When call SaveAs, it will auto download.
jExc.SaveAs("test");

```
---

Have fun !
