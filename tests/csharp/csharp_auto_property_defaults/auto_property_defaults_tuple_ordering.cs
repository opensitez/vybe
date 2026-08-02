// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
var tuple = (left: 65, right: 66); __Check((tuple.left < tuple.right).ToString(), "True");
