// vybe-test: csharp/csharp_collection_types/sorted_set_maintains_unique_sorted_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

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

var s=new System.Collections.Generic.SortedSet<int>{3,1,4,1,5};
__P((s.Count).ToString());
__P((s.Min).ToString()); __P((s.Max).ToString());
__Check("4\n1\n5");
