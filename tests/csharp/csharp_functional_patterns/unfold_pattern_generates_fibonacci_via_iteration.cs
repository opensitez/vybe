// vybe-test: csharp/csharp_functional_patterns/unfold_pattern_generates_fibonacci_via_iteration
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

System.Collections.Generic.IEnumerable<int> Fibs(){
    int a=0,b=1;
    while(true){yield return a; (a,b)=(b,a+b);}
}
var first8=Fibs().Take(8).ToArray();
Console.WriteLine(first8[7]);
