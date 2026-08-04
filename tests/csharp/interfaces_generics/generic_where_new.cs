// vybe-test: csharp/interfaces_generics/generic_where_new
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

class Factory<T> where T : new() {
    public T Create() { return new T(); }
}
class Item {
    public string Name = "default";
}
var f = new Factory<Item>();
var item = f.Create();
__P((item.Name).ToString());
__Check("default");
