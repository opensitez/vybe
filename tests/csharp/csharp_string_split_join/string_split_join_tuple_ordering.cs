// vybe-test: csharp/csharp_string_split_join/string_split_join_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
var tuple = (left: 21, right: 22); __Check((tuple.left < tuple.right).ToString(), "True");
