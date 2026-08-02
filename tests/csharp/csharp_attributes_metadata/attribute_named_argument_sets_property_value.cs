// vybe-test: csharp/csharp_attributes_metadata/attribute_named_argument_sets_property_value
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public int Priority { get; set; } public LabelAttribute(string name) { Name = name; } } [Label("job", Priority = 3)] class TaskItem { } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(TaskItem), typeof(LabelAttribute)); __Check((attr.Name).ToString(), "job"); __Check((attr.Priority).ToString(), "3");
