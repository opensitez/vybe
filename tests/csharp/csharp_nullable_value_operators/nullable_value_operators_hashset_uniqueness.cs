// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
var set = new System.Collections.Generic.HashSet<int>(); set.Add(57); set.Add(57); __Check((set.Count == 1).ToString(), "True");
