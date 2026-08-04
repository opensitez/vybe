// vybe-test: csharp/csharp_functional_patterns/method_chaining_builds_fluent_pipeline
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

var result=new[]{5,3,8,1,4}
    .Where(x=>x>2)
    .OrderBy(x=>x)
    .Select(x=>x*10)
    .First();
__P((result).ToString());
__Check("30");
