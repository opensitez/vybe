// vybe-test: csharp/csharp_delegate_variance/func_covariant_reassign_source
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<object> wide=null; System.Func<string> narrow=()=>"rebind"; wide=narrow; __Check((wide()).ToString(), "rebind");
