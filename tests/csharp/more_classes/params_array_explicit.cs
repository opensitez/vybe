// vybe-test: csharp/more_classes/params_array_explicit
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

class Program {
            static int Sum(params int[] numbers) {
                var total = 0;
                for (var i = 0; i < 5; i++) {
                    total = total + numbers[i];
                }
                return total;
            }
        }
        var arr = new int[] {1, 2, 3, 4, 5};
        __P((Program.Sum(arr)).ToString());
__Check("15");
