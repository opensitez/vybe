// vybe-test: csharp/csharp_generic_variance2/contravariant_interface_allows_base_where_derived_expected
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

using static __Harness;

IWriter<string> w=new ObjectWriter();
w.Write("hi");
__Check("hi");

interface IWriter<in T>{void Write(T v);}

class ObjectWriter:IWriter<object>{public void Write(object v)=>__P((v).ToString());}

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
