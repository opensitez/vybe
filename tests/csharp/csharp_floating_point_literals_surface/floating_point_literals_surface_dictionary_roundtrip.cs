// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

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

// floating_point_literals_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[16] = 17; __P((map.ContainsKey(16) && map[16] == 17).ToString());
__Check("True");
