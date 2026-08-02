// vybe-test: csharp/csharp_delegate_variance/func_covariant_invoke_twice_same_result
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> f=()=>"same"; System.Func<object> g=f; __Check((g()).ToString(), "same"); __Check((g()).ToString(), "same");
