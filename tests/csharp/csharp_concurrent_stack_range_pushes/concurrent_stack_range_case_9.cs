// vybe-test: csharp/csharp_concurrent_stack_range_pushes/concurrent_stack_range_case_9

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

var cs = new System.Collections.Concurrent.ConcurrentStack<int>();
cs.PushRange(new int[] { 9, 10, 11 });
int[] buf = new int[2];
int popped = cs.TryPopRange(buf);
__P(popped.ToString());
__P(buf[0].ToString());
__P(buf[1].ToString());
__Check("2\n11\n10");
