// vybe-test: csharp/csharp_char_operations/string_from_char_array_roundtrips_via_tochar_array
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = new string(new char[]{'h','i'}); __Check((s).ToString(), "hi");
