// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
var tuple = (left: 35, right: 36); __Check((tuple.left < tuple.right).ToString(), "True");
