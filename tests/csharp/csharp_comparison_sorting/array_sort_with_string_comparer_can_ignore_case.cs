// vybe-test: csharp/csharp_comparison_sorting/array_sort_with_string_comparer_can_ignore_case
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

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

var values = new[] { "b", "A", "c" }; System.Array.Sort(values, System.StringComparer.OrdinalIgnoreCase); foreach (var value in values) __P((value).ToString());
__Check("A\nb\nc");
