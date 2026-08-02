// vybe-test: csharp/csharp_generics_constraints/generic_method_can_use_list_of_t_argument
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; int Count<T>(List<T> items) { return items.Count; } __Check((Count(new List<string> { "a", "b" })).ToString(), "2");
