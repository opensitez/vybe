// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_recursive_default_calls_itself
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRec{int Fact(int n)=>n<=1?1:n*Fact(n-1);} class Math:IRec{} __Check((new Math().Fact(5)).ToString(), "120");
