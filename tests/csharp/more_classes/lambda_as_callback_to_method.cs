// vybe-test: csharp/more_classes/lambda_as_callback_to_method
// origin: languages/csharp/tests/csharp/test_more_classes.rs

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

class Util {
            public int Apply(int x) {
                return x * 2;
            }
        }
        var u = new Util();
        __P((u.Apply(21)).ToString());
__Check("42");
