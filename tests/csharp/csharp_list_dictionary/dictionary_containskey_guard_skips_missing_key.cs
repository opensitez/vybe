// vybe-test: csharp/csharp_list_dictionary/dictionary_containskey_guard_skips_missing_key
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var map = new Dictionary<string, int>(); if (map.ContainsKey("missing")) Console.WriteLine(map["missing"]); else Console.WriteLine("absent");
