// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_readonly
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

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

readonly record struct Size(int W,int H); var s=new Size(2,3); var t=s with{H=8}; __P((s.H).ToString()); __P((t.H).ToString());
__Check("3\n8");
