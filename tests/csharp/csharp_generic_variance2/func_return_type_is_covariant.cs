// vybe-test: csharp/csharp_generic_variance2/func_return_type_is_covariant
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> getStr=()=>"hello";
System.Func<object> getObj=getStr;
__Check((getObj()).ToString(), "hello");
