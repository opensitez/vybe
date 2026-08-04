// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

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

// for_loop_bounds
var map = new System.Collections.Generic.Dictionary<int, int>(); map[45] = 46; __P((map.ContainsKey(45) && map[45] == 46).ToString());
__Check("True");
