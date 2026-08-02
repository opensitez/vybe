// vybe-test: csharp/csharp_delegate_variance/func_string_to_object_covariant_return_invokes
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> getString=()=>"covariant"; System.Func<object> getObject=getString; __Check((getObject()).ToString(), "covariant");
