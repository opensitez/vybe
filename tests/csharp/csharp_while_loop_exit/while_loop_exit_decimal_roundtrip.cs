// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
