# JExcel
Build Excel via Javascript

---

This is a very simple script without rich function. I will add more function if not busy. If you are interesting in it, we can do it together.

##Example

var jExc = new JExcel();

```js
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
