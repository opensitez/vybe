// vybe-test: csharp/csharp_nested_control_flow/foreach_iteration_variable_is_fresh_each_iteration
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

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

int last = -1;
foreach (var value in new[] { 1, 2, 3 }) {
    last = value;
}
__P((last).ToString());
__Check("3");
