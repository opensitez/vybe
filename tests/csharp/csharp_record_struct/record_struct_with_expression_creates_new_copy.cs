// vybe-test: csharp/csharp_record_struct/record_struct_with_expression_creates_new_copy
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

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

record struct Point(int X,int Y);
var a=new Point(1,2);
var b=a with{X=99};
__P((a.X).ToString()); __P((b.X).ToString());
__Check("1\n99");
