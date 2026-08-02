// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
var map = new System.Collections.Generic.Dictionary<int, int>(); map[31] = 32; __Check((map.ContainsKey(31) && map[31] == 32).ToString(), "True");
