// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
var tuple = (left: 26, right: 27); __Check((tuple.left < tuple.right).ToString(), "True");
