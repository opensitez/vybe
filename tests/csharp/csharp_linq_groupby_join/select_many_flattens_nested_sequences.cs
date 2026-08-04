// vybe-test: csharp/csharp_linq_groupby_join/select_many_flattens_nested_sequences
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

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

var nested = new[] { new[]{1,2}, new[]{3,4} };
var flat = nested.SelectMany(x => x);
int sum = 0;
foreach (var n in flat) sum += n;
__P((sum).ToString());
__Check("10");
