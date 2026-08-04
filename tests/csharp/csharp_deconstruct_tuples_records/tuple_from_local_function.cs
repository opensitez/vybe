// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_from_local_function
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

System.ValueTuple<int,int> Twice(int n)=>(n,n); var (a,b)=Twice(6); __P((a*b).ToString());
__Check("36");
