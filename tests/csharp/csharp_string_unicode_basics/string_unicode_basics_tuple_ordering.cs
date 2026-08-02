// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
var tuple = (left: 19, right: 20); __Check((tuple.left < tuple.right).ToString(), "True");
