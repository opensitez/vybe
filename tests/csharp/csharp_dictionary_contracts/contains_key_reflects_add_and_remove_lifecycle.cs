// vybe-test: csharp/csharp_dictionary_contracts/contains_key_reflects_add_and_remove_lifecycle
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

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
var map = new Dictionary<int, string>();
map[1] = "one";
__P((map.ContainsKey(1) ? "Y" : "N").ToString());
map.Remove(1);
__P((map.ContainsKey(1) ? "Y" : "N").ToString());
__Check("Y\nN");
