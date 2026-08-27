// vybe-test: csharp/csharp_exceptions_hierarchy/exception_data_dictionary_stores_arbitrary_key_value
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

using static __Harness;

string r="";
try{
    var ex=new System.Exception("test");
    ex.Data["userId"]=42;
    throw ex;
}
catch(System.Exception ex){r=ex.Data["userId"].ToString();}
__P((r).ToString());
__Check("42");

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
