// vybe-test: csharp/csharp_list_dictionary/list_indexof_zero_after_head_insert
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

using System.Collections.Generic; var list = new List<int> { 2 }; list.Insert(0, 1); __P((list.IndexOf(1)).ToString());
__Check("0");
