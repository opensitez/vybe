// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_find_returns_first_of_duplicates
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

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(2); var node = ll.Find(2); __P((node.Value).ToString()); __P((node.Next.Value).ToString());
__Check("2\n2");
