// vybe-test: csharp/csharp_generic_variance2/covariant_interface_allows_derived_where_base_expected
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

using static __Harness;

IReader<object> r=new StringReader();
__P((r.Read()).ToString());
__Check("hello");

interface IReader<out T>{T Read();}

class StringReader:IReader<string>{public string Read()=>"hello";}

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
