// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

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

// string_unicode_basics
var tuple = (left: 19, right: 20); __P((tuple.left < tuple.right).ToString());
__Check("True");
