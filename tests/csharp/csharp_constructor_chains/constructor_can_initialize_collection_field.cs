// vybe-test: csharp/csharp_constructor_chains/constructor_can_initialize_collection_field
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; class Box { List<int> values; public Box() { values = new List<int> { 1, 2, 3 }; } public int Count() { return values.Count; } } __Check((new Box().Count()).ToString(), "3");
