// vybe-test: csharp/csharp_functional_patterns/function_composition_applies_in_sequence
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

System.Func<int,int> triple=x=>x*3;
System.Func<int,int> addOne=x=>x+1;
var composed=new[]{1,2,3}.Select(triple).Select(addOne);
foreach(var n in composed) Console.WriteLine(n);
