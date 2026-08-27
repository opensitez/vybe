// vybe-test: csharp/csharp_interfaces_advanced/class_implements_multiple_interfaces_with_different_methods
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

using static __Harness;

var b=new Buffer();
((IWrite)b).Write("data");
__P((((IRead)b).Read()).ToString());
__Check("data");

interface IRead{string Read();}

interface IWrite{void Write(string v);}

class Buffer:IRead,IWrite{
    string _val="";
    public string Read()=>_val;
    public void Write(string v){_val=v;}
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
