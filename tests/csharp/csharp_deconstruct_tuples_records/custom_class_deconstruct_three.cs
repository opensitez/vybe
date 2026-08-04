// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_class_deconstruct_three
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

class Box{public int A,B,C; public void Deconstruct(out int a,out int b,out int c){a=A;b=B;c=C;}} var (a,b,c)=new Box{A=1,B=2,C=3}; __P((a+b+c).ToString());
__Check("6");
