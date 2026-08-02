// vybe-test: csharp/csharp_string_methods/compare_ordinal_distinguishes_case_when_flag_false
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int r = string.Compare("A","a",System.StringComparison.Ordinal);
__Check((r < 0 || r > 0).ToString(), "True");
