// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
var values = new System.Collections.Generic.List<int> { 51, 52, 51 }; __Check((values.Count == 3).ToString(), "True");
