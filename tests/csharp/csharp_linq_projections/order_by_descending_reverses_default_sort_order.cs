// vybe-test: csharp/csharp_linq_projections/order_by_descending_reverses_default_sort_order
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

var result = new[]{3,1,4,1,5}.OrderByDescending(x => x).Distinct();
foreach(var n in result) __P((n).ToString());
__Check("5\n4\n3\n1");
