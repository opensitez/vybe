// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_in_if_else_selects_missing_branch
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using System.Collections.Generic; var map = new Dictionary<string, int>(); if (map.TryGetValue("lost", out int v)) Console.WriteLine("yes"); else Console.WriteLine("no");
