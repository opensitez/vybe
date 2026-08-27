// vybe-test: csharp/csharp_recursive_algorithms/recursive_fibonacci_returns_correct_nth_number
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

using static __Harness;

int Fib(int n)=>n<=1?n:Fib(n-1)+Fib(n-2);
__P((Fib(8)).ToString());
__Check("21");

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
