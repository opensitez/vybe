// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

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

// indexer_get_set
var map = new System.Collections.Generic.Dictionary<int, int>(); map[66] = 67; __P((map.ContainsKey(66) && map[66] == 67).ToString());
__Check("True");
