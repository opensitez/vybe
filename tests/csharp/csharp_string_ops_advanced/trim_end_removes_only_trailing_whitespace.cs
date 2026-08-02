// vybe-test: csharp/csharp_string_ops_advanced/trim_end_removes_only_trailing_whitespace
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("  hi  ".TrimEnd()).ToString(), "  hi");
