// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_first_key_is_minimum
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

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

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [10] = "ten", [2] = "two", [7] = "seven" }; int first = 0; foreach (var k in sd.Keys) { first = k; break; } __P((first).ToString());
__Check("2");
