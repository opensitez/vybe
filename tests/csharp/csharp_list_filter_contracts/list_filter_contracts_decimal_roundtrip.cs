// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
