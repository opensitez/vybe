// vybe-test: csharp/collections_advanced/list_findall
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

var list = new List<int> { 1, 2, 3, 4, 5, 6 };
var evens = list.FindAll(x => x % 2 == 0);
foreach (var x in evens) __P((x).ToString());
__Check("2\n4\n6");
