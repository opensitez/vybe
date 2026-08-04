// vybe-test: csharp/csharp_tuples_advanced/tuple_in_linq_select_creates_anonymous_projection
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

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

var items = new[]{"apple","kiwi","pear"};
var proj = items.Select(s => (Name: s, Len: s.Length));
foreach(var x in proj) __P((x.Len).ToString());
__Check("5\n4\n4");
