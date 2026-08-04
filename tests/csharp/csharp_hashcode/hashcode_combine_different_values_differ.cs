// vybe-test: csharp/csharp_hashcode/hashcode_combine_different_values_differ
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

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

int h1=System.HashCode.Combine(1,2);
int h2=System.HashCode.Combine(2,1);
__P((h1!=h2).ToString());
__Check("True");
