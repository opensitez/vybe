// vybe-test: csharp/csharp_functional_patterns/partial_application_creates_specialized_function
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

System.Func<int,System.Func<int,int>> add=a=>b=>a+b;
var add10=add(10);
__P((add10(5)).ToString());
__P((add10(20)).ToString());
__Check("15\n30");
