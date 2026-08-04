// vybe-test: csharp/csharp_tuples_advanced/tuple_returned_from_method_and_destructured_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

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

(int Min, int Max) Bounds(int[] arr) =>
    (arr.Min(), arr.Max());
var (lo, hi) = Bounds(new[]{5,1,9,3});
__P((lo).ToString()); __P((hi).ToString());
__Check("1\n9");
