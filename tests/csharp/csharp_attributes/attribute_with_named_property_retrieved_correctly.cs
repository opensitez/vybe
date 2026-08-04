// vybe-test: csharp/csharp_attributes/attribute_with_named_property_retrieved_correctly
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

[System.AttributeUsage(System.AttributeTargets.Method)]
class PriorityAttribute:System.Attribute{public int Level{get;set;}}
class Work{
    [Priority(Level=3)]
    public void DoIt(){}
}
var mi=typeof(Work).GetMethod("DoIt");
var attr=(PriorityAttribute)mi.GetCustomAttributes(typeof(PriorityAttribute),false)[0];
__P((attr.Level).ToString());
__Check("3");
