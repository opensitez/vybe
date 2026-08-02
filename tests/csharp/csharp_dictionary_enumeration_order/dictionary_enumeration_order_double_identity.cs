// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
double seed = 35; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
