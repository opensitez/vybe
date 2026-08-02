// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
int? maybe = null; int fallback = maybe ?? 31; __Check((fallback == 31).ToString(), "True");
