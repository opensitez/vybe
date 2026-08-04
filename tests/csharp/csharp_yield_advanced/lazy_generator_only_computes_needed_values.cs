// vybe-test: csharp/csharp_yield_advanced/lazy_generator_only_computes_needed_values
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

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

int calls=0;
System.Collections.Generic.IEnumerable<int> Expensive(){
    for(int i=0;;i++){calls++;yield return i;}
}
var first3=Expensive().Take(3).ToList();
__P((calls).ToString()); __P((first3[2]).ToString());
__Check("3\n2");
