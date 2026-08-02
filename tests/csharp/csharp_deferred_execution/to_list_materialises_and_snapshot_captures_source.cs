// vybe-test: csharp/csharp_deferred_execution/to_list_materialises_and_snapshot_captures_source
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var source=new System.Collections.Generic.List<int>{1,2,3};
var snapshot=source.ToList();
source.Add(4);
__Check((snapshot.Count).ToString(), "3");
