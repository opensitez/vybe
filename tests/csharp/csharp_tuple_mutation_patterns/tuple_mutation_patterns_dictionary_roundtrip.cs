// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

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

// tuple_mutation_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[37] = 38; __P((map.ContainsKey(37) && map[37] == 38).ToString());
__Check("True");
