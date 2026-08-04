// vybe-test: csharp/csharp_linq_numeric/max_by_returns_whole_element_not_just_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

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

var words=new[]{"cat","elephant","ox"};
__P((words.MaxBy(w=>w.Length)).ToString());
__Check("elephant");
