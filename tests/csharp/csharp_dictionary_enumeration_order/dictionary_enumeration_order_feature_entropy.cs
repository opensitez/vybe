// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
string feature = "dictionary_enumeration_order:35"; __Check((feature.Length >= 1).ToString(), "True");
