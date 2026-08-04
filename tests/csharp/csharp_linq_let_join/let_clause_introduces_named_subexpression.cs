// vybe-test: csharp/csharp_linq_let_join/let_clause_introduces_named_subexpression
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

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

var result =
    from s in new[]{"hello","hi","world"}
    let len=s.Length
    where len>3
    select s;
foreach(var x in result) __P((x).ToString());
__Check("hello\nworld");
