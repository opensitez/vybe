// vybe-test: csharp/csharp_map_set_collections/hashset_union_with_merges_unique_values
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var left = new HashSet<int> { 1, 2 }; left.UnionWith(new[] { 2, 3 }); foreach (var item in left) Console.WriteLine(item);
