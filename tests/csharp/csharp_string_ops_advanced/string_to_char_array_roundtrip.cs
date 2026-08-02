// vybe-test: csharp/csharp_string_ops_advanced/string_to_char_array_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] chars="abc".ToCharArray();
__Check((new string(chars)).ToString(), "abc");
