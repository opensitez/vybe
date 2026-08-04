// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_to_existing_locals
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

class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var target=new Pair{A=5,B=6}; int x,y; (x,y)=target; __P((x+y).ToString());
__Check("11");
