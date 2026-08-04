// vybe-test: csharp/csharp_static_type_behaviors/static_dictionary_state_is_mutated_by_each_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

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

using System.Collections.Generic;
class Tracker {
    static Dictionary<string, int> counts = new Dictionary<string, int>();
    public void Hit(string key) {
        if (!counts.ContainsKey(key)) counts[key] = 0;
        counts[key]++;
    }
    public static int Read(string key) { return counts[key]; }
}
var a = new Tracker();
var b = new Tracker();
a.Hit("api");
b.Hit("api");
__P((Tracker.Read("api")).ToString());
__Check("2");
