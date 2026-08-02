// vybe-test: csharp/csharp_attributes_metadata/attribute_on_enum_is_readable_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Enum)] class GroupAttribute : Attribute { public string Name { get; } public GroupAttribute(string name) { Name = name; } } [Group("status")] enum State { Idle } var attr = (GroupAttribute)Attribute.GetCustomAttribute(typeof(State), typeof(GroupAttribute)); __Check((attr.Name).ToString(), "status");
