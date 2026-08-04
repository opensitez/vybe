// vybe-test: csharp/csharp_generics/generic_stack_queue
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

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

var stack = new Stack<string>();
stack.Push("first");
stack.Push("second");
__P((stack.Pop()).ToString());
var queue = new Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
__P((queue.Dequeue()).ToString());
__Check("second\n1");
