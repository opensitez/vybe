// vybe-test: csharp/csharp_dictionary_operations/foreach_over_dictionary_yields_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<int,int>{{1,10}};
foreach(var pair in d) __P((pair.Key + ":" + pair.Value).ToString());
__Check("1:10");
