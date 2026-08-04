// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_then_increment
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

int slot = 2;
System.Threading.Interlocked.Exchange(ref slot, 10);
__P((System.Threading.Interlocked.Increment(ref slot)).ToString());
__P((slot).ToString());
__Check("11\n11");
