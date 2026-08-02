// vybe-test: csharp/csharp_string_ops_advanced/trim_start_removes_only_leading_whitespace
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("  hi  ".TrimStart()).ToString(), "hi  ");
