// vybe-test: csharp/csharp_random_random/seeded_random_produces_deterministic_sequence
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

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

var r1=new System.Random(99); var r2=new System.Random(99);
__P((r1.Next()==r2.Next()).ToString());
__Check("True");
