// vybe-test: csharp/collections/list_add_numbers
// origin: languages/csharp/tests/csharp/test_collections.rs

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

var list = new List<int>();
        list.Add(10);
        list.Add(20);
        list.Add(30);
        var sum = 0;
        foreach (var x in list) { sum = sum + x; }
        __P((sum).ToString());
__Check("60");
