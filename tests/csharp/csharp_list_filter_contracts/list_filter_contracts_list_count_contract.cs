// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
var values = new System.Collections.Generic.List<int> { 31, 32, 31 }; __Check((values.Count == 3).ToString(), "True");
