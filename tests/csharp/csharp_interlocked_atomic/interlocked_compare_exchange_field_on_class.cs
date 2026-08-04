// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

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
    public int Value = 0;
    public int Cas(int expected, int desired) {
        return System.Threading.Interlocked.CompareExchange(ref Value, desired, expected);
    }
}
var c = new Counter();
__P((c.Cas(0, 11)).ToString());
__P((c.Value).ToString());
__Check("0\n11");
