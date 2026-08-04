// vybe-test: csharp/csharp_record_struct/record_struct_deconstruct_in_let_statement
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

record struct Vec(int X,int Y);
var v=new Vec(3,4);
var(x,y)=v;
__P((x+y).ToString());
__Check("7");
