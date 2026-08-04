// vybe-test: csharp/csharp_list_dictionary/list_nested_foreach_counts_each_inner_length
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

using System.Collections.Generic; var outer = new List<List<int>> { new List<int> { 1, 2 }, new List<int> { 3 } }; foreach (var inner in outer) __P((inner.Count).ToString());
__Check("2\n1");
