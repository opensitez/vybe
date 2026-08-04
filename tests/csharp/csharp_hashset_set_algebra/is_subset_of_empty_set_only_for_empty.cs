// vybe-test: csharp/csharp_hashset_set_algebra/is_subset_of_empty_set_only_for_empty
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

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

using System.Collections.Generic; var empty = new HashSet<int>(); var nonempty = new HashSet<int> { 1 }; __P((empty.IsSubsetOf(nonempty)).ToString()); __P((nonempty.IsSubsetOf(empty)).ToString());
__Check("True\nFalse");
