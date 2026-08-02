// vybe-test: csharp/csharp_dynamic/dynamic_variable_reassigned_to_different_type
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

dynamic v=42;
__Check((v.GetType().Name).ToString(), "Int32");
v="hello";
__Check((v.GetType().Name).ToString(), "String");
