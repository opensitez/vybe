// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_one_equals_increment
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

int a = 6;
int b = 6;
System.Threading.Interlocked.Increment(ref a);
System.Threading.Interlocked.Add(ref b, 1);
__P((a + b).ToString());
__Check("14");
