# Excel Wrapper
The cells works like array, when: A=0, B=1, C=2..., etc.
The rows init on 1, like excel.

#Create  
`$excel = new Excel('COMPANY NAME');`  

#Add text to cell  
setValue($cell, $row, $value)  
`$excel->setValue(1, 1, "Event");`  
Add word "Event" on B:1

#Merge
`$excel->merge(1,1, 3, true);`  

#Background Cell
`$excel->backgroundCell("#fff000", 1,1);`  

#Download
`$excel->download('fileName');`
#Save
If you need store excel data in a variable
`$data = $excel->save();`

#Full Example
`$excel = new Excel('Eycopia');
 $excel->merge(1,1, 3, true);
 $excel->setValue(1, 1, "Event");
 $excel->backgroundCell("#fff000", 1,1);
 $excel->setValue(0, 2, "Days", true);
 $excel->setValue(1, 2, "Monday", true);
 $excel->setValue(2, 2, "Tuesday", true);
 $excel->setValue(3, 2, "Wednesday", true);
 
 $excel->setValue(0, 3, "Hackaton", true)
     ->setValue(1, 3, "12", true)
     ->setValue(2, 3, "34", true)
     ->setValue(3, 3, "15", true);
 
 $excel->setValue(0,4, "Networking", true)
     ->setValue(1,4, "10", true)
     ->setValue(2,4, "12", true)
     ->setValue(3,4, "43", true);
 
 $excel->download('juan');`
