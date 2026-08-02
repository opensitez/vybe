// vybe-test: csharp/csharp_lambdas/lambda_with_list_foreach
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

using System.Collections.Generic;
var items = new List<int>();
items.Add(1);
items.Add(2);
items.Add(3);
items.ForEach(x => Console.WriteLine(x));
