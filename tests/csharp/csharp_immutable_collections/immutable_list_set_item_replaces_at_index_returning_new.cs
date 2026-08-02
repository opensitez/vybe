// vybe-test: csharp/csharp_immutable_collections/immutable_list_set_item_replaces_at_index_returning_new
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list=System.Collections.Immutable.ImmutableList.Create(1,2,3);
var updated=list.SetItem(1,99);
__Check((list[1]).ToString(), "2"); __Check((updated[1]).ToString(), "99");
