// vybe-test: csharp/csharp_delegate_variance/func_string_array_to_object_array_covariant
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string[]> getStrings=()=>new string[]{"a"}; System.Func<object[]> getObjects=getStrings; __Check((getObjects()[0]).ToString(), "a");
