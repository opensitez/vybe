// vybe-test: csharp/csharp_console_write/writeline_bool_capitalises_true
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// console_write
__Check((true).ToString(), "True");
