// vybe-test: csharp/csharp_with_expression_records/with_nominal_chain
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

record C{public int A{get;init;} public int B{get;init;}} var e=((new C{A=1,B=2}) with{A=3}) with{B=4}; __P((e.A).ToString()); __P((e.B).ToString());
__Check("3\n4");
