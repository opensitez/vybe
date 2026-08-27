// vybe-test: csharp/csharp_collections_immutable_dictionary_ops/immutable_dict_case_7

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var d1 = System.Collections.Immutable.ImmutableDictionary<string, int>.Empty;
var d2 = d1.Add("k_7", 7);
__P(d1.Count.ToString());
__P(d2["k_7"].ToString());
__Check("0\n7");
