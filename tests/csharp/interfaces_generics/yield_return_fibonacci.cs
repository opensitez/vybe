// vybe-test: csharp/interfaces_generics/yield_return_fibonacci
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

class Fib {
    public static IEnumerable<int> Sequence(int count) {
        int a = 0, b = 1;
        for (int i = 0; i < count; i++) {
            yield return a;
            int temp = a + b;
            a = b;
            b = temp;
        }
    }
}
foreach (var n in Fib.Sequence(8)) __P((n).ToString());
__Check("0\n1\n1\n2\n3\n5\n8\n13");
