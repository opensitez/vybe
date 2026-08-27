// vybe-test: csharp/csharp_recursive_algorithms/mutual_recursion_even_odd_check
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

using static __Harness;

bool IsEven(int n){if(n==0)return true; return IsOdd(n-1);}
bool IsOdd(int n){if(n==0)return false; return IsEven(n-1);}
__P((IsEven(4)).ToString());
__P((IsOdd(3)).ToString());
__Check("True\nTrue");

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
