// vybe-test: csharp/csharp_delegate_types/removing_handler_from_multicast_leaves_remaining
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

using static __Harness;

int count = 0;
System.Action a = () => count++;
System.Action b = () => count++;
System.Action multi = a;
multi += b;
multi -= a;
multi();
__P((count).ToString());
__Check("1");

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
