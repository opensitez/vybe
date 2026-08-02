// vybe-test: csharp/csharp_bcl_collections/linked_list_add_first_inserts_before_current_head
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.LinkedList<int>();
list.AddLast(2);
list.AddFirst(1);
__Check((list.First.Value).ToString(), "1");
