// vybe-test: csharp/csharp_primary_constructors/primary_constructor_list_param_count
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag(System.Collections.Generic.List<int> items) { public int Count => items.Count; }
__Check((new Bag(new System.Collections.Generic.List<int> { 1, 2 }).Count).ToString(), "2");
