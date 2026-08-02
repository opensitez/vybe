// vybe-test: csharp/csharp_functional_patterns/partial_application_creates_specialized_function
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,System.Func<int,int>> add=a=>b=>a+b;
var add10=add(10);
__Check((add10(5)).ToString(), "15");
__Check((add10(20)).ToString(), "30");
