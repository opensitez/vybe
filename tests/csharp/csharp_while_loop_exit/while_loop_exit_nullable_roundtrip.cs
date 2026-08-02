// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
int? maybe = 47; __Check((maybe.HasValue && maybe.Value == 47).ToString(), "True");
