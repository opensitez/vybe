// vybe-test: csharp/csharp_comparison_sorting/order_by_with_key_projection_sorts_by_length
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

using System.Linq; var values = new[] { "bbb", "a", "cc" }.OrderBy(text => text.Length); foreach (var value in values) __P((value).ToString());
__Check("a\ncc\nbbb");
