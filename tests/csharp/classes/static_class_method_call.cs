// vybe-test: csharp/classes/static_class_method_call
// origin: languages/csharp/tests/csharp/test_classes.rs

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

class MathHelper {
            public static int Square(int x) { return x * x; }
            public static int Double(int x) { return x * 2; }
        }
        __P((MathHelper.Square(5)).ToString());
        __P((MathHelper.Double(7)).ToString());
__Check("25\n14");
