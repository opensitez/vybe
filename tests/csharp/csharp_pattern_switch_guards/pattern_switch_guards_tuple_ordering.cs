// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
var tuple = (left: 42, right: 43); __Check((tuple.left < tuple.right).ToString(), "True");
