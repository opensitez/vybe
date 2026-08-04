// vybe-test: csharp/csharp_map_set_collections/dictionary_remove_erases_entry_from_map
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

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

using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 9 }; map.Remove("x"); __P((map.ContainsKey("x")).ToString());
__Check("False");
