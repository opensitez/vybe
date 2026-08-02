// vybe-test: csharp/csharp_conversion_methods/convert_change_type_dynamically_converts_to_target_type
// origin: languages/csharp/tests/csharp/test_csharp_conversion_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object result=System.Convert.ChangeType("42",typeof(int));
__Check((result).ToString(), "42");
