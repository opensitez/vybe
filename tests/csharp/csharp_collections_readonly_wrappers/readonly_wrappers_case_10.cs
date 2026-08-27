// vybe-test: csharp/csharp_collections_readonly_wrappers/readonly_wrappers_case_10

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

var raw = new System.Collections.Generic.List<int>() { 10, 11 };
var roc = new System.Collections.ObjectModel.ReadOnlyCollection<int>(raw);
__P(roc.Count.ToString());
__P(roc[0].ToString());
__Check("2\n10");
