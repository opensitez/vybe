// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
var set = new System.Collections.Generic.HashSet<int>(); set.Add(19); set.Add(19); __Check((set.Count == 1).ToString(), "True");
