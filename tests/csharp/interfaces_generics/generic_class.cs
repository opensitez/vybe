// vybe-test: csharp/interfaces_generics/generic_class
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

class Box<T> {
    public T Value;
    public Box(T val) { Value = val; }
}
var intBox = new Box<int>(42);
var strBox = new Box<string>("hello");
__P((intBox.Value).ToString());
__P((strBox.Value).ToString());
__Check("42\nhello");
