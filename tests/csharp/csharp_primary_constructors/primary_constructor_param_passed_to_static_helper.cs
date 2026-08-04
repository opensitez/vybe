// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_passed_to_static_helper
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Worker(int id) {
    static string Format(int value) => "id=" + value;
    public string Show() => Format(id);
}
__P((new Worker(3).Show()).ToString());
__Check("id=3");
