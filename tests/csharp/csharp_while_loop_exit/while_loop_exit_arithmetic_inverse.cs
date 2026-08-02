// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
int seed = 47; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
