// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_record_twice
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

record Pair(int A,int B); var p=new Pair(1,2); var (a,b)=p; var (c,d)=p; __P((a+c).ToString()); __P((b+d).ToString());
__Check("2\n4");
