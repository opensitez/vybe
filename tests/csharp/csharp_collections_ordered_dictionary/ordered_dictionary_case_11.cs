// vybe-test: csharp/csharp_collections_ordered_dictionary/ordered_dictionary_case_11

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

var od = new System.Collections.Specialized.OrderedDictionary();
od.Add("key_11", "val_11");
__P(od[0]?.ToString());
__Check("val_11");
