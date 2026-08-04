// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_remove_deletes_entry
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

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a", [2] = "b" }; sd.Remove(1); __P((sd.Count).ToString());
__Check("1");
