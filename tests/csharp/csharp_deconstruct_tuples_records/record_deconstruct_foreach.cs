// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_foreach
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

record Point(int X,int Y); var pts=new[]{new Point(1,2),new Point(3,4)}; int sum=0; foreach(var (x,y) in pts) sum+=x+y; __P((sum).ToString());
__Check("10");
