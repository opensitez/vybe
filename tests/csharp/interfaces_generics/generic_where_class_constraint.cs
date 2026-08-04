// vybe-test: csharp/interfaces_generics/generic_where_class_constraint
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

class Container<T> where T : class {
    public T Value;
    public bool IsNull() { return Value == null; }
}
var c = new Container<string>();
__P((c.IsNull()).ToString());
c.Value = "hello";
__P((c.IsNull()).ToString());
__Check("True\nFalse");
