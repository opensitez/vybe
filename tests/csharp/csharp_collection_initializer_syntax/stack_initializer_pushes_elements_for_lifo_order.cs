// vybe-test: csharp/csharp_collection_initializer_syntax/stack_initializer_pushes_elements_for_lifo_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;
using System.Collections.Generic;

var stack = new Stack<int>();
stack.Push(1);
stack.Push(2);
__P((stack.Pop()).ToString());
__Check("2");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
