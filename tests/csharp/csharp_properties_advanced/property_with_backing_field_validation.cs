// vybe-test: csharp/csharp_properties_advanced/property_with_backing_field_validation
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

using static __Harness;

var a=new Age{Value=25}
;
__P((a.Value).ToString());
__Check("25");

class Age{
    int _value;
    public int Value{
        get=>_value;
        set{if(value<0)throw new System.ArgumentException();_value=value;}
    }
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
