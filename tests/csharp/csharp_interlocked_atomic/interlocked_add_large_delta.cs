// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_large_delta
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

int total = 100;
__P((System.Threading.Interlocked.Add(ref total, 900)).ToString());
__P((total).ToString());
__Check("1000\n1000");
