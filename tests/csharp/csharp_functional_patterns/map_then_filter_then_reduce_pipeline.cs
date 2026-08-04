// vybe-test: csharp/csharp_functional_patterns/map_then_filter_then_reduce_pipeline
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

var result=new[]{1,2,3,4,5}
    .Select(x=>x*x)
    .Where(x=>x>5)
    .Sum();
__P((result).ToString());
__Check("50");
