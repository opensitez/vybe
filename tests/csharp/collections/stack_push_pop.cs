// vybe-test: csharp/collections/stack_push_pop
// origin: languages/csharp/tests/csharp/test_collections.rs

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

var s = new Stack<int>();
        s.Push(1);
        s.Push(2);
        s.Push(3);
        __P((s.Pop()).ToString());
        __P((s.Pop()).ToString());
__Check("3\n2");
