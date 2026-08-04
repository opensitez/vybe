// vybe-test: csharp/csharp_extension_methods/extension_method_on_list_can_report_item_count
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

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

using Demo; using System.Collections.Generic; namespace Demo { public static class ListExt { public static string Describe<T>(this List<T> values) { return "count=" + values.Count; } } } __P((new List<int> { 1, 2 }.Describe()).ToString());
__Check("count=2");
