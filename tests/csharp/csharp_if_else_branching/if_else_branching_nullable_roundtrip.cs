// vybe-test: csharp/csharp_if_else_branching/if_else_branching_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
int? maybe = 44; __Check((maybe.HasValue && maybe.Value == 44).ToString(), "True");
