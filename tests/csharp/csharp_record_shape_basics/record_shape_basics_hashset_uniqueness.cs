// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
var set = new System.Collections.Generic.HashSet<int>(); set.Add(39); set.Add(39); __Check((set.Count == 1).ToString(), "True");
