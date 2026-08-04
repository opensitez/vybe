// vybe-test: csharp/csharp_checked_context_math/checked_context_math_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

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

// checked_context_math
var tuple = (left: 12, right: 13); __P((tuple.left < tuple.right).ToString());
__Check("True");
