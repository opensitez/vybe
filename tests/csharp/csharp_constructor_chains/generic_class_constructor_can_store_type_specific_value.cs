// vybe-test: csharp/csharp_constructor_chains/generic_class_constructor_can_store_type_specific_value
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

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

class Box<T> { T value; public Box(T value) { this.value = value; } public T Read() { return value; } } __P((new Box<string>("text").Read()).ToString());
__Check("text");
