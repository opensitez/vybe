// vybe-test: csharp/csharp_map_set_collections/stack_push_and_pop_follow_lifo_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

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

using System.Collections.Generic; var stack = new Stack<int>(); stack.Push(1); stack.Push(2); __P((stack.Pop()).ToString()); __P((stack.Pop()).ToString());
__Check("2\n1");
