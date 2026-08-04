// vybe-test: csharp/csharp_map_set_collections/sorted_set_enumerates_numbers_in_ascending_order
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

using System.Collections.Generic; var set = new SortedSet<int> { 5, 1, 3 }; foreach (var item in set) __P((item).ToString());
__Check("1\n3\n5");
