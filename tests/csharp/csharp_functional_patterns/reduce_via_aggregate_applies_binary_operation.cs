// vybe-test: csharp/csharp_functional_patterns/reduce_via_aggregate_applies_binary_operation
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

var product=new[]{1,2,3,4,5}.Aggregate((acc,x)=>acc*x);
__P((product).ToString());
__Check("120");
