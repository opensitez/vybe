// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_string_elements_maintain_insertion_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<string>(); q.Enqueue("a"); q.Enqueue("b"); __P((q.Dequeue()).ToString()); __P((q.Dequeue()).ToString());
__Check("a\nb");
