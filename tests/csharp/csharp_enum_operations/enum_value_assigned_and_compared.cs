// vybe-test: csharp/csharp_enum_operations/enum_value_assigned_and_compared
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

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

enum Color{Red,Green,Blue} var c=Color.Green; __P((c==Color.Green).ToString());
__Check("True");
