// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_in_if_else_selects_found_branch
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using System.Collections.Generic; var map = new Dictionary<string, int> { ["found"] = 11 }; if (map.TryGetValue("found", out int v)) Console.WriteLine("yes:" + v); else Console.WriteLine("no");
