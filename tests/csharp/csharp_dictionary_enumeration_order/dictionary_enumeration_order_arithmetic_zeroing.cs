// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
int seed = 35; __Check((seed - seed == 0).ToString(), "True");
