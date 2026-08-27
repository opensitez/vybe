// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

using static __Harness;

// operator_overloading_matrix
int sum = 1;
int n = 0;
while (n < 4) { sum += 1; n += 1; }
__P((sum == 5).ToString());
__Check("True");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
