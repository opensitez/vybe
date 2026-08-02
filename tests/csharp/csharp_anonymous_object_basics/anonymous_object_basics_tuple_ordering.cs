// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
var tuple = (left: 38, right: 39); __Check((tuple.left < tuple.right).ToString(), "True");
