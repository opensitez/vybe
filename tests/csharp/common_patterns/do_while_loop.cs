// vybe-test: csharp/common_patterns/do_while_loop
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

int x = 1;
do {
    __P((x).ToString());
    x *= 3;
} while (x < 100);
__Check("1\n3\n9\n27\n81");
