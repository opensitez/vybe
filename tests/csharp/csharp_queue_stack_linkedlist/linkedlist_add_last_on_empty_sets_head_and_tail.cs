// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_last_on_empty_sets_head_and_tail
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

using System.Collections.Generic; var ll = new LinkedList<string>(); ll.AddLast("solo"); __P((ll.First.Value).ToString()); __P((ll.Last.Value).ToString());
__Check("solo\nsolo");
