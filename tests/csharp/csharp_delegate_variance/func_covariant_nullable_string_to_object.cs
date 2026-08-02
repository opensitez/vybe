// vybe-test: csharp/csharp_delegate_variance/func_covariant_nullable_string_to_object
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> f=()=>null; System.Func<object> g=f; __Check((g()==null).ToString(), "True");
