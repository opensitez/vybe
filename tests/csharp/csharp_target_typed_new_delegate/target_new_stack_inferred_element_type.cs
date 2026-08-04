// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_stack_inferred_element_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

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

System.Collections.Generic.Stack<int> s = new();
s.Push(1); s.Push(2);
__P((s.Pop()).ToString()); __P((s.Pop()).ToString());
__Check("2\n1");
