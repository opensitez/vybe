// vybe-test: csharp/csharp_lambda_loop_capture_semantics/foreach_iteration_lambda_sees_each_elements_value_not_final_index
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
foreach (var value in new[] { 10, 20, 30 }) {
    actions.Add(() => value);
}
foreach (var run in actions) __P((run()).ToString());
__Check("10\n20\n30");
