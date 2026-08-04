// vybe-test: csharp/csharp_if_else_branching/if_else_branching_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

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

// if_else_branching
var map = new System.Collections.Generic.Dictionary<int, int>(); map[44] = 45; __P((map.ContainsKey(44) && map[44] == 45).ToString());
__Check("True");
