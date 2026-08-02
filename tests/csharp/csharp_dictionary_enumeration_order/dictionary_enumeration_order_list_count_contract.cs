// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
var values = new System.Collections.Generic.List<int> { 35, 36, 35 }; __Check((values.Count == 3).ToString(), "True");
