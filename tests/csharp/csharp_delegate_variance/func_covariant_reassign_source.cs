// vybe-test: csharp/csharp_delegate_variance/func_covariant_reassign_source
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

using static __Harness;

Func<DerivedClass> getDerived = () => new DerivedClass();
Func<BaseClass> getBase = getDerived;
BaseClass b = getBase();
__P((b != null).ToString());
__Check("True");

class BaseClass { }
class DerivedClass : BaseClass { }
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
