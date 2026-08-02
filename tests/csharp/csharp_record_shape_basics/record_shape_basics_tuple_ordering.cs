// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
var tuple = (left: 39, right: 40); __Check((tuple.left < tuple.right).ToString(), "True");
