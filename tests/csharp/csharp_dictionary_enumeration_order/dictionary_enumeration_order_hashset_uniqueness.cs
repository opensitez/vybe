// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_enumeration_order
var set = new System.Collections.Generic.HashSet<int>(); set.Add(35); set.Add(35); __Check((set.Count == 1).ToString(), "True");
