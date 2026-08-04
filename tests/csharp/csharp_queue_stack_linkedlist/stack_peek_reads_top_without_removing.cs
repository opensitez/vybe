// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_peek_reads_top_without_removing
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

using System.Collections.Generic; var s = new Stack<int>(); s.Push(4); s.Push(5); __P((s.Peek()).ToString()); __P((s.Count).ToString());
__Check("5\n2");
