use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(dictionary_try_get_value_reads_existing_entry, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 3 }; if (map.TryGetValue("a", out var value)) Console.WriteLine(value);"#, ["3"]);
csharp_case!(dictionary_try_get_value_reports_missing_key, r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); Console.WriteLine(map.TryGetValue("a", out var value));"#, ["False"]);
csharp_case!(dictionary_contains_key_detects_present_member, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 9 }; Console.WriteLine(map.ContainsKey("x"));"#, ["True"]);
csharp_case!(dictionary_remove_erases_entry_from_map, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 9 }; map.Remove("x"); Console.WriteLine(map.ContainsKey("x"));"#, ["False"]);
csharp_case!(dictionary_iteration_exposes_key_value_pairs, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) Console.WriteLine(pair.Key + ":" + pair.Value);"#, ["b:2", "a:1"]);
csharp_case!(sorted_dictionary_enumerates_keys_in_sorted_order, r#"using System.Collections.Generic; var map = new SortedDictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) Console.WriteLine(pair.Key + ":" + pair.Value);"#, ["a:1", "b:2"]);
csharp_case!(hashset_rejects_duplicate_value_addition, r#"using System.Collections.Generic; var set = new HashSet<int>(); Console.WriteLine(set.Add(3)); Console.WriteLine(set.Add(3));"#, ["True", "False"]);
csharp_case!(hashset_contains_reports_membership, r#"using System.Collections.Generic; var set = new HashSet<string> { "alpha", "beta" }; Console.WriteLine(set.Contains("beta"));"#, ["True"]);
csharp_case!(hashset_union_with_merges_unique_values, r#"using System.Collections.Generic; var left = new HashSet<int> { 1, 2 }; left.UnionWith(new[] { 2, 3 }); foreach (var item in left) Console.WriteLine(item);"#, ["1", "2", "3"]);
csharp_case!(hashset_intersect_with_keeps_shared_values_only, r#"using System.Collections.Generic; var left = new HashSet<int> { 1, 2, 3 }; left.IntersectWith(new[] { 2, 3, 4 }); foreach (var item in left) Console.WriteLine(item);"#, ["2", "3"]);
csharp_case!(sorted_set_enumerates_numbers_in_ascending_order, r#"using System.Collections.Generic; var set = new SortedSet<int> { 5, 1, 3 }; foreach (var item in set) Console.WriteLine(item);"#, ["1", "3", "5"]);
csharp_case!(linked_list_add_first_and_add_last_preserve_order, r#"using System.Collections.Generic; var items = new LinkedList<string>(); items.AddFirst("middle"); items.AddFirst("start"); items.AddLast("end"); foreach (var item in items) Console.WriteLine(item);"#, ["start", "middle", "end"]);
csharp_case!(queue_enqueue_and_dequeue_follow_fifo_order, r#"using System.Collections.Generic; var queue = new Queue<int>(); queue.Enqueue(1); queue.Enqueue(2); Console.WriteLine(queue.Dequeue()); Console.WriteLine(queue.Dequeue());"#, ["1", "2"]);
csharp_case!(stack_push_and_pop_follow_lifo_order, r#"using System.Collections.Generic; var stack = new Stack<int>(); stack.Push(1); stack.Push(2); Console.WriteLine(stack.Pop()); Console.WriteLine(stack.Pop());"#, ["2", "1"]);
csharp_case!(dictionary_indexer_updates_existing_value, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map["a"] = 4; Console.WriteLine(map["a"]);"#, ["4"]);
csharp_case!(dictionary_values_collection_reports_count, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; Console.WriteLine(map.Values.Count);"#, ["2"]);
csharp_case!(dictionary_keys_collection_can_be_enumerated, r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; foreach (var key in map.Keys) Console.WriteLine(key);"#, ["a", "b"]);
csharp_case!(hashset_remove_erases_existing_member, r#"using System.Collections.Generic; var set = new HashSet<int> { 1, 2 }; set.Remove(1); Console.WriteLine(set.Contains(1));"#, ["False"]);
csharp_case!(sorted_dictionary_indexer_retrieves_inserted_value, r#"using System.Collections.Generic; var map = new SortedDictionary<int, string>(); map[2] = "two"; Console.WriteLine(map[2]);"#, ["two"]);
csharp_case!(linked_list_find_returns_matching_node_value, r#"using System.Collections.Generic; var items = new LinkedList<string>(); items.AddLast("a"); items.AddLast("b"); var node = items.Find("b"); Console.WriteLine(node.Value);"#, ["b"]);