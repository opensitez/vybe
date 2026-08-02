// vybe-test: csharp/csharp_delegate_variance/func_bool_to_object_covariant_return_invokes
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<bool> getBool=()=>true; System.Func<object> getObject=getBool; __Check((getObject()).ToString(), "True");
