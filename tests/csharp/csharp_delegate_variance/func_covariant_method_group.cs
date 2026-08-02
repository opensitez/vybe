// vybe-test: csharp/csharp_delegate_variance/func_covariant_method_group
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static string Make()=>"group"; System.Func<string> narrow=Make; System.Func<object> wide=narrow; __Check((wide()).ToString(), "group");
