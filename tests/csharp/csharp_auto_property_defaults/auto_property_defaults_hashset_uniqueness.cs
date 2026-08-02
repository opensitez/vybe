// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
var set = new System.Collections.Generic.HashSet<int>(); set.Add(65); set.Add(65); __Check((set.Count == 1).ToString(), "True");
