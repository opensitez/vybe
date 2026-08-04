// vybe-test: csharp/csharp_ref_out_in/ref_parameter_mutates_caller_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

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

void Double(ref int x){x*=2;}
int n=5; Double(ref n); __P((n).ToString());
__Check("10");
