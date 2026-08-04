// vybe-test: csharp/csharp_linq_advanced/min_by_returns_element_with_minimum_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

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

var words=new[]{"a","bbb","cc"};
__P((words.MinBy(w=>w.Length)).ToString());
__Check("a");
