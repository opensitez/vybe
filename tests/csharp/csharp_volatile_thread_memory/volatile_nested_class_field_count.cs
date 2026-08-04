// vybe-test: csharp/csharp_volatile_thread_memory/volatile_nested_class_field_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

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

class Outer {
    public class Inner {
        public volatile int Value = 0;
    }
}
var inner = new Outer.Inner();
inner.Value = 13;
__P((inner.Value).ToString());
__Check("13");
