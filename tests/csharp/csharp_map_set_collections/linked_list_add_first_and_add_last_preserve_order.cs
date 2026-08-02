// vybe-test: csharp/csharp_map_set_collections/linked_list_add_first_and_add_last_preserve_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var items = new LinkedList<string>(); items.AddFirst("middle"); items.AddFirst("start"); items.AddLast("end"); foreach (var item in items) Console.WriteLine(item);
