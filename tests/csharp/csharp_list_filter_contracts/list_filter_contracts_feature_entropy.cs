// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
string feature = "list_filter_contracts:31"; __Check((feature.Length >= 1).ToString(), "True");
