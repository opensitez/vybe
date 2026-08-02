// vybe-test: csharp/csharp_map_set_collections/hashset_intersect_with_keeps_shared_values_only
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var left = new HashSet<int> { 1, 2, 3 }; left.IntersectWith(new[] { 2, 3, 4 }); foreach (var item in left) Console.WriteLine(item);
