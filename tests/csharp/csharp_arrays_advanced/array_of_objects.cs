// vybe-test: csharp/csharp_arrays_advanced/array_of_objects
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

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

class Item {
    public string Name;
    public Item(string n) { Name = n; }
}
var items = new[] { new Item("a"), new Item("b"), new Item("c") };
foreach (var item in items) {
    __P((item.Name).ToString());
}
__Check("a\nb\nc");
