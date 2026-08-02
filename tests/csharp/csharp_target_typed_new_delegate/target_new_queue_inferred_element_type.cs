// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_queue_inferred_element_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Queue<int> q = new();
q.Enqueue(4); q.Enqueue(5);
__Check((q.Dequeue()).ToString(), "4"); __Check((q.Dequeue()).ToString(), "5");
