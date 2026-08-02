// vybe-test: csharp/csharp_if_else_branching/if_else_branching_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
string feature = "if_else_branching"; __Check((feature.Length > 0).ToString(), "True");
