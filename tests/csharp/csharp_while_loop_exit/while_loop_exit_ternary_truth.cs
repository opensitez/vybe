// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
int seed = 47; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
