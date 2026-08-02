// vybe-test: csharp/csharp_immutable_collections/immutable_list_remove_returns_new_list
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list=System.Collections.Immutable.ImmutableList.Create(1,2,3);
var smaller=list.Remove(2);
__Check((list.Count).ToString(), "3"); __Check((smaller.Count).ToString(), "2");
