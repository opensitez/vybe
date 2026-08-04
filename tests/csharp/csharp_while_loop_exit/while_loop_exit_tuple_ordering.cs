// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

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

// while_loop_exit
var tuple = (left: 47, right: 48); __P((tuple.left < tuple.right).ToString());
__Check("True");
