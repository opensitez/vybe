// vybe-test: csharp/csharp_map_set_collections/linked_list_find_returns_matching_node_value
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var items = new LinkedList<string>(); items.AddLast("a"); items.AddLast("b"); var node = items.Find("b"); __Check((node.Value).ToString(), "b");
