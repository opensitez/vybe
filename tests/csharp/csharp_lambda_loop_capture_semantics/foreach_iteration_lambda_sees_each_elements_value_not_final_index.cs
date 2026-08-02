// vybe-test: csharp/csharp_lambda_loop_capture_semantics/foreach_iteration_lambda_sees_each_elements_value_not_final_index
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
foreach (var value in new[] { 10, 20, 30 }) {
    actions.Add(() => value);
}
foreach (var run in actions) Console.WriteLine(run());
