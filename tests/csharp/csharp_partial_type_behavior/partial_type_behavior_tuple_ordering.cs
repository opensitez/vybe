// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
var tuple = (left: 70, right: 71); __Check((tuple.left < tuple.right).ToString(), "True");
