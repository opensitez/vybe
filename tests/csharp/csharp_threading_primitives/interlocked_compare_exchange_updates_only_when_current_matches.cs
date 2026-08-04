// vybe-test: csharp/csharp_threading_primitives/interlocked_compare_exchange_updates_only_when_current_matches
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

int slot = 7;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 99, 7);
__P((previous).ToString());
__P((slot).ToString());
__Check("7\n99");
