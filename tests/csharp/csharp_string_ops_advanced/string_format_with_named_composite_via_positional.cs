// vybe-test: csharp/csharp_string_ops_advanced/string_format_with_named_composite_via_positional
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0:000}", 7)).ToString(), "007");
