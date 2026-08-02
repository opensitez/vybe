// vybe-test: csharp/csharp_immutable_collections/immutable_list_add_returns_new_list_old_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.Collections.Immutable.ImmutableList<int>.Empty;
var b=a.Add(1).Add(2).Add(3);
__Check((a.Count).ToString(), "0"); __Check((b.Count).ToString(), "3");
