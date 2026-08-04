// vybe-test: csharp/linq_lambdas/linq_orderbydescending
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

var nums = new List<int> { 3, 1, 4, 1, 5 };
var sorted = nums.OrderByDescending(x => x).ToList();
foreach (var x in sorted) __P((x).ToString());
__Check("5\n4\n3\n1\n1");
