// vybe-test: csharp/csharp_generic_collections/immutable_list_add_returns_new_list_without_mutating_original
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var original = System.Collections.Immutable.ImmutableList.Create(1,2,3);
var extended = original.Add(4);
__Check((original.Count).ToString(), "3");
__Check((extended.Count).ToString(), "4");
