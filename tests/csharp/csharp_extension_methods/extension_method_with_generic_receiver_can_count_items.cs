// vybe-test: csharp/csharp_extension_methods/extension_method_with_generic_receiver_can_count_items
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

using Demo; using System.Collections.Generic; namespace Demo { public static class EnumerableExt { public static int CountItems<T>(this IEnumerable<T> items) { int total = 0; foreach (var _ in items) total++; return total; } } } __P((new[] { 1, 2, 3 }.CountItems()).ToString());
__Check("3");
