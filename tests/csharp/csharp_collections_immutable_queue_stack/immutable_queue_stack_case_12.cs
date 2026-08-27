// vybe-test: csharp/csharp_collections_immutable_queue_stack/immutable_queue_stack_case_12

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var q1 = System.Collections.Immutable.ImmutableQueue<int>.Empty;
var q2 = q1.Enqueue(12);
int val = q2.Peek();
__P(val.ToString());
__Check("12");
