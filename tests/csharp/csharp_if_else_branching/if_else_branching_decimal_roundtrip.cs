// vybe-test: csharp/csharp_if_else_branching/if_else_branching_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
