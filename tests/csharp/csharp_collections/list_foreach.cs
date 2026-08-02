// vybe-test: csharp/csharp_collections/list_foreach
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using System.Collections.Generic;
var list = new List<int>();
list.Add(1);
list.Add(2);
list.Add(3);
int sum = 0;
foreach (var x in list) {
    sum += x;
}
Console.WriteLine(sum);
