// vybe-test: csharp/csharp_attributes/attribute_with_named_property_retrieved_correctly
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

using static __Harness;

var mi=typeof(Work).GetMethod("DoIt");
var attr=(PriorityAttribute)mi.GetCustomAttributes(typeof(PriorityAttribute),false)[0];
__P((attr.Level).ToString());
__Check("3");

[System.AttributeUsage(System.AttributeTargets.Method)]
class PriorityAttribute:System.Attribute{public int Level{get;set;}}

class Work{
    [Priority(Level=3)]
    public void DoIt(){}
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
