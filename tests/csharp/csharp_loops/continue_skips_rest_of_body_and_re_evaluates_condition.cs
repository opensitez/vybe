// vybe-test: csharp/csharp_loops/continue_skips_rest_of_body_and_re_evaluates_condition
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

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

int s=0; for(int i=0;i<5;i++) { if(i%2==0) continue; s+=i; } __P((s).ToString());
__Check("4");
