// vybe-test: csharp/csharp_if_else_branching/if_else_branching_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
int seed = 44; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
