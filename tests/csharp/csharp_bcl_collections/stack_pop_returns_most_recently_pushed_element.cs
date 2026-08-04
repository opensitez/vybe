// vybe-test: csharp/csharp_bcl_collections/stack_pop_returns_most_recently_pushed_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

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

var stack = new System.Collections.Generic.Stack<int>();
stack.Push(1);
stack.Push(2);
__P((stack.Pop()).ToString());
__Check("2");
