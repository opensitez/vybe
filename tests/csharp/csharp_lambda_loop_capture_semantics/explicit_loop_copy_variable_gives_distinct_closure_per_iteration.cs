// vybe-test: csharp/csharp_lambda_loop_capture_semantics/explicit_loop_copy_variable_gives_distinct_closure_per_iteration
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

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

using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
for (int i = 0; i < 3; i++) {
    int copy = i;
    actions.Add(() => copy);
}
foreach (var run in actions) __P((run()).ToString());
__Check("0\n1\n2");
