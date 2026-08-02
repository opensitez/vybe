// vybe-test: csharp/csharp_lambda_loop_capture_semantics/explicit_loop_copy_variable_gives_distinct_closure_per_iteration
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
for (int i = 0; i < 3; i++) {
    int copy = i;
    actions.Add(() => copy);
}
foreach (var run in actions) Console.WriteLine(run());
