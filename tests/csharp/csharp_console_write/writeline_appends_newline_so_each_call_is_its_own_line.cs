// vybe-test: csharp/csharp_console_write/writeline_appends_newline_so_each_call_is_its_own_line
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// console_write
__Check(("a").ToString(), "a"); __Check(("b").ToString(), "b");
