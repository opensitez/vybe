// vybe-test: csharp/csharp_covariance_contravariance/func_return_type_covariance_allows_derived_func_in_base_func
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> getString = () => "hi";
System.Func<object> getObject = getString;
__Check((getObject()).ToString(), "hi");
