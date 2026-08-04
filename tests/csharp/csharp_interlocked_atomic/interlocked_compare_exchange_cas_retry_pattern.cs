// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_cas_retry_pattern
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

int slot = 0;
int observed = slot;
int desired = observed + 1;
while (observed != System.Threading.Interlocked.CompareExchange(ref slot, desired, observed)) {
    observed = slot;
    desired = observed + 1;
}
__P((slot).ToString());
__Check("1");
