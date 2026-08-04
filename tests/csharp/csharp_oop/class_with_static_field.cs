// vybe-test: csharp/csharp_oop/class_with_static_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

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

class Counter {
    public static int Count = 0;
    public Counter() { Count++; }
}
var a = new Counter();
var b = new Counter();
var c = new Counter();
__P((Counter.Count).ToString());
__Check("3");
