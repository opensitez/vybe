// vybe-test: csharp/csharp_bcl_collections/sorted_dictionary_enumerator_yields_keys_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

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

var map = new System.Collections.Generic.SortedDictionary<int, string>();
map[3] = "c";
map[1] = "a";
int firstKey = 0;
foreach (var pair in map) { firstKey = pair.Key; break; }
__P((firstKey).ToString());
__Check("1");
