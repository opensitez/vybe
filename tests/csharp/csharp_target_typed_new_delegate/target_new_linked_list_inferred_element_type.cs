// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_linked_list_inferred_element_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.LinkedList<int> list = new();
list.AddLast(10); list.AddLast(20);
__Check((list.First.Value).ToString(), "10"); __Check((list.Last.Value).ToString(), "20");
