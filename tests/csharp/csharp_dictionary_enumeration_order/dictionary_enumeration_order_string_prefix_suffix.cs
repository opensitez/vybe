// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
string feature = "dictionary_enumeration_order"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
