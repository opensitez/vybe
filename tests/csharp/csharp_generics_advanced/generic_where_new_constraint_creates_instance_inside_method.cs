// vybe-test: csharp/csharp_generics_advanced/generic_where_new_constraint_creates_instance_inside_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

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

class Widget { public int Val = 5; }
T Make<T>() where T : new() => new T();
__P((Make<Widget>().Val).ToString());
__Check("5");
