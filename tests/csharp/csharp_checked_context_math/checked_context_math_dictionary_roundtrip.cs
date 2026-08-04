// vybe-test: csharp/csharp_checked_context_math/checked_context_math_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

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

// checked_context_math
var map = new System.Collections.Generic.Dictionary<int, int>(); map[12] = 13; __P((map.ContainsKey(12) && map[12] == 13).ToString());
__Check("True");
