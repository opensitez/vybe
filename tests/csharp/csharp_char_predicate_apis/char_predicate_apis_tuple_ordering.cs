// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
var tuple = (left: 23, right: 24); __Check((tuple.left < tuple.right).ToString(), "True");
