// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_two
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

record struct Point(int X,int Y); var p=new Point(1,2); var q=p with{X=3,Y=4}; __P((p.Y).ToString()); __P((q.X).ToString()); __P((q.Y).ToString());
__Check("2\n3\n4");
