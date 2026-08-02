// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
var tuple = (left: 31, right: 32); __Check((tuple.left < tuple.right).ToString(), "True");
