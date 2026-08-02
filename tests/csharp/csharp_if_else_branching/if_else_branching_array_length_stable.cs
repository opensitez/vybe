// vybe-test: csharp/csharp_if_else_branching/if_else_branching_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
int seed = 44; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
