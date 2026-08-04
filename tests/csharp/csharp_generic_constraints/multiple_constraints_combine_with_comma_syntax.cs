// vybe-test: csharp/csharp_generic_constraints/multiple_constraints_combine_with_comma_syntax
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

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

interface IName { string Name(); }
T Make<T>() where T : IName, new() => new T();
class Item : IName { public string Name() => "item"; }
__P((Make<Item>().Name()).ToString());
__Check("item");
