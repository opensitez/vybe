// vybe-test: csharp/csharp_integer_arithmetic/pre_and_post_increment_on_different_variables
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

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

int left = 2; int right = 5; __P((++left + right++).ToString()); __P((left).ToString()); __P((right).ToString());
__Check("8\n3\n6");
