// vybe-test: csharp/csharp_collections_immutable_dictionary_ops/immutable_dict_case_11

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
var d2 = d1.Add("k_11", 11);
__P(d1.Count.ToString());
__P(d2["k_11"].ToString());
__Check("0\n11");
