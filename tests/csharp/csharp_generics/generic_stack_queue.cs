// vybe-test: csharp/csharp_generics/generic_stack_queue
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var stack = new Stack<string>();
stack.Push("first");
stack.Push("second");
__Check((stack.Pop()).ToString(), "second");
var queue = new Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
__Check((queue.Dequeue()).ToString(), "1");
