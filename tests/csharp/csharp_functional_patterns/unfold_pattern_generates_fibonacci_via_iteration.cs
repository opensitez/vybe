// vybe-test: csharp/csharp_functional_patterns/unfold_pattern_generates_fibonacci_via_iteration
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

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

System.Collections.Generic.IEnumerable<int> Fibs(){
    int a=0,b=1;
    while(true){yield return a; (a,b)=(b,a+b);}
}
var first8=Fibs().Take(8).ToArray();
__P((first8[7]).ToString());
__Check("13");
