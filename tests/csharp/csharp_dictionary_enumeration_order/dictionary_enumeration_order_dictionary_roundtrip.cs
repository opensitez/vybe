// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
var map = new System.Collections.Generic.Dictionary<int, int>(); map[35] = 36; __Check((map.ContainsKey(35) && map[35] == 36).ToString(), "True");
