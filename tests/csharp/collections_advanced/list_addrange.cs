// vybe-test: csharp/collections_advanced/list_addrange
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

var list = new List<int> { 1, 2, 3 };
list.AddRange(new int[] { 4, 5 });
__P((list.Count).ToString());
foreach (var x in list) __P((x).ToString());
__Check("5\n1\n2\n3\n4\n5");
