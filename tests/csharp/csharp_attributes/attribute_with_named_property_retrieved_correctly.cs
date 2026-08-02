// vybe-test: csharp/csharp_attributes/attribute_with_named_property_retrieved_correctly
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((attr.Level).ToString(), "3");
