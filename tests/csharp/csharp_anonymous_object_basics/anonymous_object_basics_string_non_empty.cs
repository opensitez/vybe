// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

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

// anonymous_object_basics
string feature = "anonymous_object_basics"; __P((feature.Length > 0).ToString());
__Check("True");
