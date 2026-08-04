// vybe-test: csharp/csharp_attributes_metadata/attribute_named_argument_sets_property_value
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

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

using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public int Priority { get; set; } public LabelAttribute(string name) { Name = name; } } [Label("job", Priority = 3)] class TaskItem { } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(TaskItem), typeof(LabelAttribute)); __P((attr.Name).ToString()); __P((attr.Priority).ToString());
__Check("job\n3");
