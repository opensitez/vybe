// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_contains_false_for_absent_value
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

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); __P((ll.Contains(2)).ToString());
__Check("False");
