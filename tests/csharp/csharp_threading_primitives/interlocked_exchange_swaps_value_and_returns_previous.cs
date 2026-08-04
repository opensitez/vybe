// vybe-test: csharp/csharp_threading_primitives/interlocked_exchange_swaps_value_and_returns_previous
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

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

int slot = 1;
__P((System.Threading.Interlocked.Exchange(ref slot, 9)).ToString());
__P((slot).ToString());
__Check("1\n9");
