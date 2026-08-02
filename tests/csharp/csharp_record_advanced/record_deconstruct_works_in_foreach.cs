// vybe-test: csharp/csharp_record_advanced/record_deconstruct_works_in_foreach
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

record Point(int X,int Y);
var pts=new[]{new Point(1,2),new Point(3,4)};
int sumX=0;
foreach(var(x,_) in pts) sumX+=x;
Console.WriteLine(sumX);
