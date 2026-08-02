// vybe-test: csharp/csharp_if_else_branching/if_else_branching_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
var tuple = (left: 44, right: 45); __Check((tuple.left < tuple.right).ToString(), "True");
