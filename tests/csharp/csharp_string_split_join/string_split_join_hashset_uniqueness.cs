// vybe-test: csharp/csharp_string_split_join/string_split_join_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

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

// string_split_join
var set = new System.Collections.Generic.HashSet<int>(); set.Add(21); set.Add(21); __P((set.Count == 1).ToString());
__Check("True");
