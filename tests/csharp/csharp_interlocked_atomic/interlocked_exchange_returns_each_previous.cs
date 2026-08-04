// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_returns_each_previous
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

int slot = 4;
int old1 = System.Threading.Interlocked.Exchange(ref slot, 5);
int old2 = System.Threading.Interlocked.Exchange(ref slot, 6);
__P((old1 + old2).ToString());
__P((slot).ToString());
__Check("9\n6");
