// vybe-test: csharp/csharp_comparison_sorting/array_sort_with_parallel_keys_reorders_items_by_key
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

var keys = new[] { 2, 1 }; var items = new[] { "b", "a" }; System.Array.Sort(keys, items); foreach (var value in items) __P((value).ToString());
__Check("a\nb");
