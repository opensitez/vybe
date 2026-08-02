// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
var tuple = (left: 20, right: 21); __Check((tuple.left < tuple.right).ToString(), "True");
