// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_nullable_record
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

record Maybe(int? N); var (n)=new Maybe(5); __P((n).ToString());
__Check("5");
