// vybe-test: csharp/csharp_array_apis/array_clone_creates_independent_shallow_copy
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

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

var source = new[] { 1, 2 }; var clone = (int[])source.Clone(); clone[0] = 9; __P((source[0]).ToString()); __P((clone[0]).ToString());
__Check("1\n9");
