// vybe-test: csharp/linq_lambdas/linq_skip_take
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

var nums = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
var page = nums.Skip(2).Take(3).ToList();
foreach (var x in page) __P((x).ToString());
__Check("3\n4\n5");
