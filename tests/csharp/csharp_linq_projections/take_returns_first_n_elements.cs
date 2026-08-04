// vybe-test: csharp/csharp_linq_projections/take_returns_first_n_elements
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

var result = new[]{10,20,30,40}.Take(2);
foreach(var n in result) __P((n).ToString());
__Check("10\n20");
