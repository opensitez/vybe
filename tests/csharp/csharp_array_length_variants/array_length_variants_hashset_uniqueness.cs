// vybe-test: csharp/csharp_array_length_variants/array_length_variants_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
var set = new System.Collections.Generic.HashSet<int>(); set.Add(25); set.Add(25); __Check((set.Count == 1).ToString(), "True");
