// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
