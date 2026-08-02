// vybe-test: csharp/csharp_delegate_variance/func_covariant_multiline_lambda
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> narrow=()=>{return "multi";}; System.Func<object> wide=narrow; __Check((wide()).ToString(), "multi");
