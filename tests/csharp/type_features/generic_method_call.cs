// vybe-test: csharp/type_features/generic_method_call
// origin: languages/csharp/tests/csharp/test_type_features.rs

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

var list = new List<int>();
        list.Add(1);
        list.Add(2);
        list.Add(3);
        __P((list.Count).ToString());
__Check("3");
