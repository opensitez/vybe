// vybe-test: csharp/collections_advanced/list_convertall
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

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

var nums = new List<int> { 1, 2, 3 };
var strings = nums.ConvertAll(x => x.ToString());
foreach (var s in strings) __P((s).ToString());
__Check("1\n2\n3");
