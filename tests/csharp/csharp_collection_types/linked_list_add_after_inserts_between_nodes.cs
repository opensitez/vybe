// vybe-test: csharp/csharp_collection_types/linked_list_add_after_inserts_between_nodes
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ll=new System.Collections.Generic.LinkedList<int>();
var n1=ll.AddFirst(1);
ll.AddAfter(n1,3);
ll.AddAfter(n1,2);
__Check((ll.First.Next.Value).ToString(), "2");
