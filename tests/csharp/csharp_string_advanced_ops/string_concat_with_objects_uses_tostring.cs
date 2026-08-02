// vybe-test: csharp/csharp_string_advanced_ops/string_concat_with_objects_uses_tostring
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Concat("val=",42," ok=",true)).ToString(), "val=42 ok=True");
