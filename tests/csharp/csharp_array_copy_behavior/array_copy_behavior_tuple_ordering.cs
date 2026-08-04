// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

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

// array_copy_behavior
var tuple = (left: 26, right: 27); __P((tuple.left < tuple.right).ToString());
__Check("True");
