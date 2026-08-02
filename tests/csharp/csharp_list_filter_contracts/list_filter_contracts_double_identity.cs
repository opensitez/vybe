// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
double seed = 31; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
