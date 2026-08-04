// vybe-test: csharp/linq_lambdas/linq_where_select_tolist
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
var result = nums.Where(x => x % 2 == 0).Select(x => x * 10).ToList();
foreach (var x in result) __P((x).ToString());
__Check("20\n40\n60\n80");
