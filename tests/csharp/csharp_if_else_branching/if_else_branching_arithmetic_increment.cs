// vybe-test: csharp/csharp_if_else_branching/if_else_branching_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
int seed = 44; __Check((seed + 1 > seed).ToString(), "True");
