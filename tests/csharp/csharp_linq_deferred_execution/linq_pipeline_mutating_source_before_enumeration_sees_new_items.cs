// vybe-test: csharp/csharp_linq_deferred_execution/linq_pipeline_mutating_source_before_enumeration_sees_new_items
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Collections.Generic;
using System.Linq;
var data = new List<int> { 1, 2 };
var query = data.Select(x => x * 10);
data.Add(3);
foreach (var value in query) Console.WriteLine(value);
