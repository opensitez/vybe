// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

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

// string_builder_usage
var set = new System.Collections.Generic.HashSet<int>(); set.Add(20); set.Add(20); __P((set.Count == 1).ToString());
__Check("True");
