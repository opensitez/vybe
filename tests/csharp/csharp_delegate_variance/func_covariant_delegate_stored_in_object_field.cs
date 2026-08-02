// vybe-test: csharp/csharp_delegate_variance/func_covariant_delegate_stored_in_object_field
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> f=()=>"field"; System.Func<object> g=f; object boxed=g; __Check((((System.Func<object>)boxed)()).ToString(), "field");
