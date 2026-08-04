// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_to_locals
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

record R(int A,int B); var r=new R(2,3); int x,y; (x,y)=r; __P((x*y).ToString());
__Check("6");
