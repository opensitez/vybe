// vybe-test: csharp/common_patterns/fibonacci_iterative
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

int n = 10;
int a = 0, b = 1;
for (int i = 0; i < n; i++) {
    __P((a).ToString());
    int tmp = a + b;
    a = b;
    b = tmp;
}
__Check("0\n1\n1\n2\n3\n5\n8\n13\n21\n34");
