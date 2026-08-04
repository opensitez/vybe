// vybe-test: csharp/linq_runtime/linq_select_strings
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

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

var list = new List<string>();
list.Add("hello"); list.Add("world");
list.Select(s => s.ToUpper()).ForEach(s => __P((s).ToString()));
__Check("HELLO\nWORLD");
