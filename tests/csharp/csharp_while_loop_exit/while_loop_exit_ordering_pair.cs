// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
int seed = 47; int right = seed + 1; __Check((seed < right).ToString(), "True");
