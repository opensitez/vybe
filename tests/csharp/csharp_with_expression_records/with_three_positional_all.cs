// vybe-test: csharp/csharp_with_expression_records/with_three_positional_all
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

record Triple(int A,int B,int C); var u=(new Triple(1,2,3)) with{A=4,B=5,C=6}; __P((u.A+u.B+u.C).ToString());
__Check("15");
