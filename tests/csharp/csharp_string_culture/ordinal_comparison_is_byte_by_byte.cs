// vybe-test: csharp/csharp_string_culture/ordinal_comparison_is_byte_by_byte
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int r=string.CompareOrdinal("a","A");
__Check((r>0).ToString(), "True");
