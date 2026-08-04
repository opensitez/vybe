// vybe-test: csharp/csharp_sorted_collections/sorted_set_subset_of_larger_sorted_set
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

using System.Collections.Generic; var small = new SortedSet<int> { 2, 3 }; var big = new SortedSet<int> { 1, 2, 3, 4 }; __P((small.IsSubsetOf(big)).ToString());
__Check("True");
