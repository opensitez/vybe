// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_filter_contracts
string feature = "list_filter_contracts"; __Check((feature[0] == feature[0]).ToString(), "True");
