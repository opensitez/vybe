// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_struct_deconstruct
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

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

struct Pair{public int X,Y; public void Deconstruct(out int x,out int y){x=X;y=Y;}} var (x,y)=new Pair{X=4,Y=6}; __P((x*y).ToString());
__Check("24");
