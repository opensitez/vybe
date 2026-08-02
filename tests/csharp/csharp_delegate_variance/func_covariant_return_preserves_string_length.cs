// vybe-test: csharp/csharp_delegate_variance/func_covariant_return_preserves_string_length
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> src=()=>"abcd"; System.Func<object> widened=src; __Check((((string)widened()).Length).ToString(), "4");
