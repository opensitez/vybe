// vybe-test: csharp/csharp_lambda_loop_capture_semantics/for_loop_lambda_captures_shared_counter_showing_last_value_at_invoke_time
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
for (int i = 0; i < 3; i++) {
    actions.Add(() => i);
}
foreach (var run in actions) Console.WriteLine(run());
