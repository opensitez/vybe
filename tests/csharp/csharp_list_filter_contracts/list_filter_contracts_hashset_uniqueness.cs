// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
var set = new System.Collections.Generic.HashSet<int>(); set.Add(31); set.Add(31); __Check((set.Count == 1).ToString(), "True");
