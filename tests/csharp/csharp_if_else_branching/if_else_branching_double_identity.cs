// vybe-test: csharp/csharp_if_else_branching/if_else_branching_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
double seed = 44; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
