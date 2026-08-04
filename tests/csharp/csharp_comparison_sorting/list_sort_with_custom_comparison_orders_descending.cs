// vybe-test: csharp/csharp_comparison_sorting/list_sort_with_custom_comparison_orders_descending
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

using System.Collections.Generic; var list = new List<int> { 1, 3, 2 }; list.Sort((left, right) => right.CompareTo(left)); foreach (var value in list) __P((value).ToString());
__Check("3\n2\n1");
