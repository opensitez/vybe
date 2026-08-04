// vybe-test: csharp/csharp_linq_let_join/order_by_in_query_syntax_sorts_ascending
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

var q=from n in new[]{3,1,2} orderby n select n;
foreach(var x in q) __P((x).ToString());
__Check("1\n2\n3");
