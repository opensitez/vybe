// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_factory_method_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

using static __Harness;

string res = GenericHelper.Call<StaticImpl>();
__P(res);
__Check("Static_static_abstract_factory_method_on_interface");

interface IStaticProvider<TSelf> where TSelf : IStaticProvider<TSelf> {
    static abstract string GetValue();
}
class StaticImpl : IStaticProvider<StaticImpl> {
    public static string GetValue() => "Static_static_abstract_factory_method_on_interface";
}
class GenericHelper {
    public static string Call<T>() where T : IStaticProvider<T> => T.GetValue();
}
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
