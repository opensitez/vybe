// vybe-test: csharp/csharp_with_expression_records/with_two_branches_same_source
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

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

record Pair(int A,int B); var p=new Pair(1,2); var x=p with{A=9}; var y=p with{B=8}; __P((x.A).ToString()); __P((y.B).ToString());
__Check("9\n8");
