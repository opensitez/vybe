// vybe-test: csharp/csharp_hashset_set_algebra/intersect_with_string_names_keeps_common
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

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

using System.Collections.Generic; var a = new HashSet<string> { "x", "y" }; a.IntersectWith(new[] { "y", "z" }); __P((a.Contains("y")).ToString()); __P((a.Count).ToString());
__Check("True\n1");
