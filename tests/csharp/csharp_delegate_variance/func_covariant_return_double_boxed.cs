// vybe-test: csharp/csharp_delegate_variance/func_covariant_return_double_boxed
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<double> f=()=>3.14; System.Func<object> g=f; __Check((g()).ToString(), "3.14");
