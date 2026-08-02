// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_queue_string_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Queue<string> q = new();
q.Enqueue("first"); q.Enqueue("second");
__Check((q.Peek()).ToString(), "first");
