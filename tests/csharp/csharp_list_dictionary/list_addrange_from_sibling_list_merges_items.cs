// vybe-test: csharp/csharp_list_dictionary/list_addrange_from_sibling_list_merges_items
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

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

using System.Collections.Generic; var list = new List<int> { 1 }; var extra = new List<int> { 2, 3 }; list.AddRange(extra); __P((list.Count).ToString());
__Check("3");
