// vybe-test: csharp/csharp_hashcode/hashcode_add_produces_same_as_combine_for_two_values
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

var hc=new System.HashCode();
hc.Add(1); hc.Add(2);
int h1=hc.ToHashCode();
int h2=System.HashCode.Combine(1,2);
__P((h1==h2).ToString());
__Check("True");
