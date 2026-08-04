// vybe-test: csharp/csharp_generics_advanced/generic_pair_swaps_values_through_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

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

(T, T) Swap<T>(T a, T b) => (b, a);
var (x, y) = Swap(1, 2);
__P((x).ToString()); __P((y).ToString());
__Check("2\n1");
