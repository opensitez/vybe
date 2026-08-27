// vybe-test: csharp/csharp_generics_where/where_interface_constraint_calls_interface_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

using static __Harness;

string GetName<T>(T t) where T:IName=>t.Name();
__P((GetName(new A())).ToString());
__Check("A");

interface IName{string Name();}

class A:IName{public string Name()=>"A";}

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
