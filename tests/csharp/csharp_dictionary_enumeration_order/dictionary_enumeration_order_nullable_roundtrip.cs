// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
int? maybe = 35; __Check((maybe.HasValue && maybe.Value == 35).ToString(), "True");
