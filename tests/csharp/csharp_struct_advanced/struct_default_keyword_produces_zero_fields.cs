// vybe-test: csharp/csharp_struct_advanced/struct_default_keyword_produces_zero_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

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

struct Vec{public int X,Y,Z;}
var v=default(Vec);
__P((v.X==0&&v.Y==0&&v.Z==0).ToString());
__Check("True");
