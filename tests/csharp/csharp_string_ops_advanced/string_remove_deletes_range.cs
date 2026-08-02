// vybe-test: csharp/csharp_string_ops_advanced/string_remove_deletes_range
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello world".Remove(5,6)).ToString(), "hello");
