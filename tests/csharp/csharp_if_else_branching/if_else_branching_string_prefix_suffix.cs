// vybe-test: csharp/csharp_if_else_branching/if_else_branching_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
string feature = "if_else_branching"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
