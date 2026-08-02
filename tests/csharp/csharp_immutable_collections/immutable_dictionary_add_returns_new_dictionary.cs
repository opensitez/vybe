// vybe-test: csharp/csharp_immutable_collections/immutable_dictionary_add_returns_new_dictionary
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=System.Collections.Immutable.ImmutableDictionary<string,int>.Empty;
var d2=d.Add("a",1).Add("b",2);
__Check((d.Count).ToString(), "0"); __Check((d2["b"]).ToString(), "2");
