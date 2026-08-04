// vybe-test: csharp/csharp_extension_methods/extension_method_can_compose_with_generator_output
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

using Demo; using System.Collections.Generic; namespace Demo { public static class NumberExt { public static IEnumerable<int> Twice(this IEnumerable<int> values) { foreach (var value in values) yield return value * 2; } } } foreach (var value in new[] { 1, 2 }.Twice()) __P((value).ToString());
__Check("2\n4");
