// vybe-test: csharp/csharp_delegates_advanced/func_returns_func_partial_application
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,System.Func<int,int>> multiply=factor=>n=>n*factor;
var triple=multiply(3);
__Check((triple(7)).ToString(), "21");
