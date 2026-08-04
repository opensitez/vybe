// vybe-test: csharp/csharp_record_advanced/record_struct_has_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

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

record struct Vec(int X,int Y);
var a=new Vec(1,2); var b=a; // copy
b=b with{X=99};
__P((a.X).ToString());
__Check("1");
