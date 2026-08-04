// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

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

// dictionary_insert_lookup
var set = new System.Collections.Generic.HashSet<int>(); set.Add(34); set.Add(34); __P((set.Count == 1).ToString());
__Check("True");
