// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

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

// for_loop_bounds
var tuple = (left: 45, right: 46); __P((tuple.left < tuple.right).ToString());
__Check("True");
