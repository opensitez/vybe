// vybe-test: csharp/csharp_string_ops_advanced/substring_single_arg_returns_suffix_from_index
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello world".Substring(6)).ToString(), "world");
