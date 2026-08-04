// vybe-test: csharp/csharp_linq_projections/take_while_stops_at_first_failing_predicate
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

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

var result = new[]{1,3,5,4,7}.TakeWhile(x => x%2!=0);
foreach(var n in result) __P((n).ToString());
__Check("1\n3\n5");
