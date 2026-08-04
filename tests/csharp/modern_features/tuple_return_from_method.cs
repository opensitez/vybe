// vybe-test: csharp/modern_features/tuple_return_from_method
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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

class MathOps {
    public static (int Min, int Max) MinMax(int[] arr) {
        int min = arr[0], max = arr[0];
        foreach (var x in arr) {
            if (x < min) min = x;
            if (x > max) max = x;
        }
        return (min, max);
    }
}
var result = MathOps.MinMax(new int[] { 3, 1, 4, 1, 5, 9 });
__P((result.Min).ToString());
__P((result.Max).ToString());
__Check("1\n9");
