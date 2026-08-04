// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

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

// static_constructor_guard
var map = new System.Collections.Generic.Dictionary<int, int>(); map[69] = 70; __P((map.ContainsKey(69) && map[69] == 70).ToString());
__Check("True");
