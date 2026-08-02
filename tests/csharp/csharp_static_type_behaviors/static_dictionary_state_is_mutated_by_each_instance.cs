// vybe-test: csharp/csharp_static_type_behaviors/static_dictionary_state_is_mutated_by_each_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((Tracker.Read("api")).ToString(), "2");
