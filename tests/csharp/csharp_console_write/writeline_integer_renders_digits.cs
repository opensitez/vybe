// vybe-test: csharp/csharp_console_write/writeline_integer_renders_digits
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// console_write
__Check((42).ToString(), "42");
