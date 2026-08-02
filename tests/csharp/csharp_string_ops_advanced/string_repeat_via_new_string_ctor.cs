// vybe-test: csharp/csharp_string_ops_advanced/string_repeat_via_new_string_ctor
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new string('*', 5)).ToString(), "*****");
