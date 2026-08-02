// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
int seed = 31; int right = seed + 1; __Check((seed < right).ToString(), "True");
