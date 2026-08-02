// vybe-test: csharp/csharp_if_else_branching/if_else_branching_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
var values = new System.Collections.Generic.List<int> { 44, 45, 44 }; __Check((values.Count == 3).ToString(), "True");
