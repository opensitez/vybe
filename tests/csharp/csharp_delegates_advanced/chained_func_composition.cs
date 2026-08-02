// vybe-test: csharp/csharp_delegates_advanced/chained_func_composition
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> double_it=x=>x*2;
System.Func<int,int> add_three=x=>x+3;
System.Func<int,int> combined=x=>add_three(double_it(x));
__Check((combined(5)).ToString(), "13");
