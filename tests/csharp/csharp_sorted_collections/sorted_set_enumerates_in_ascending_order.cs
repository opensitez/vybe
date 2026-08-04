// vybe-test: csharp/csharp_sorted_collections/sorted_set_enumerates_in_ascending_order
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

using System.Collections.Generic; var ss = new SortedSet<int> { 5, 1, 3, 4, 2 }; foreach (var x in ss) __P((x).ToString());
__Check("1\n2\n3\n4\n5");
