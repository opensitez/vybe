// vybe-test: csharp/csharp_linq_materialization/linq_to_list_materializes_query_before_source_mutation
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
using System.Linq;
var source = new List<int> { 1, 2 };
var snapshot = source.Select(x => x).ToList();
source.Add(3);
__Check((snapshot.Count).ToString(), "2");
__Check((source.Count).ToString(), "3");
