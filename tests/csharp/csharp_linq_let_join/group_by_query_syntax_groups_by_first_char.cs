// vybe-test: csharp/csharp_linq_let_join/group_by_query_syntax_groups_by_first_char
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

var words=new[]{"apple","ant","banana"};
var groups=from w in words group w by w[0];
int count=0;
foreach(var g in groups) count++;
__P((count).ToString());
__Check("2");
