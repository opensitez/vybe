// vybe-test: csharp/csharp_yield_advanced/lazy_generator_only_computes_needed_values
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

int calls=0;
System.Collections.Generic.IEnumerable<int> Expensive(){
    for(int i=0;;i++){calls++;yield return i;}
}
var first3=Expensive().Take(3).ToList();
Console.WriteLine(calls); Console.WriteLine(first3[2]);
