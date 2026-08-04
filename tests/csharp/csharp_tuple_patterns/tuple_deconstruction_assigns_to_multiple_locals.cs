// vybe-test: csharp/csharp_tuple_patterns/tuple_deconstruction_assigns_to_multiple_locals
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

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

(string name,int age)=("Alice",30);
__P((name).ToString()); __P((age).ToString());
__Check("Alice\n30");
