// vybe-test: csharp/csharp_string_comparison/string_comparer_ordinal_ignore_case_used_in_sorted_set
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

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

var set = new System.Collections.Generic.SortedSet<string>(
    System.StringComparer.OrdinalIgnoreCase);
set.Add("Apple"); set.Add("apple");
__P((set.Count).ToString());
__Check("1");
