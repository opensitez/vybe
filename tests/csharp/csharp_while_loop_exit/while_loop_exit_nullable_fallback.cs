// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
int? maybe = null; int fallback = maybe ?? 47; __Check((fallback == 47).ToString(), "True");
