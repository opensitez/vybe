// vybe-test: csharp/csharp_collection_types/sorted_set_remove_eliminates_element
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

var s=new System.Collections.Generic.SortedSet<int>{1,2,3};
s.Remove(2);
__P((s.Count).ToString());
__Check("2");
