// vybe-test: csharp/csharp_constructor_patterns/parameterless_constructor_required_for_generic_new_constraint
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

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

class Widget{public int Value=7;}
T Make<T>() where T:new()=>new T();
__P((Make<Widget>().Value).ToString());
__Check("7");
