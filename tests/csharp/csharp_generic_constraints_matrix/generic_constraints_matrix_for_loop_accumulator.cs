// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

using static __Harness;

// generic_constraints_matrix
int sum = 0;
for (int i = 0; i < 3; i++) { sum += i; }
__P((sum == 3).ToString());
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
