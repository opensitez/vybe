// vybe-test: csharp/csharp_string_ops_advanced/string_insert_inserts_at_position
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("helo".Insert(3,"l")).ToString(), "hello");
