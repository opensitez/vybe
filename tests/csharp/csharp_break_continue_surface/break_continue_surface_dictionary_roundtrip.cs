// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

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

// break_continue_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[49] = 50; __P((map.ContainsKey(49) && map[49] == 50).ToString());
__Check("True");
