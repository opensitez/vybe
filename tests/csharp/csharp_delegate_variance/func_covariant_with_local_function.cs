// vybe-test: csharp/csharp_delegate_variance/func_covariant_with_local_function
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Local()=>"local"; System.Func<string> f=Local; System.Func<object> g=f; __Check((g()).ToString(), "local");
