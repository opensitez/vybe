// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
var set = new System.Collections.Generic.HashSet<int>(); set.Add(20); set.Add(20); __Check((set.Count == 1).ToString(), "True");
