// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
var tuple = (left: 69, right: 70); __Check((tuple.left < tuple.right).ToString(), "True");
