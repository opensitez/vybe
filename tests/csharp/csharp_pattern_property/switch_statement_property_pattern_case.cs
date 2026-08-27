// vybe-test: csharp/csharp_pattern_property/switch_statement_property_pattern_case
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Node{Id=5}
;
string tag="";
switch(o){case Node{Id:5}:tag="match";break;default:tag="miss";break;}
__P((tag).ToString());
__Check("match");

class Node { public int Id; }

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
