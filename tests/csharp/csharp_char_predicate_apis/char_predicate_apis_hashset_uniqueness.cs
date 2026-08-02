// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
var set = new System.Collections.Generic.HashSet<int>(); set.Add(23); set.Add(23); __Check((set.Count == 1).ToString(), "True");
