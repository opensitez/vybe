// vybe-test: csharp/csharp_attributes_metadata/attribute_get_custom_attributes_can_return_strongly_typed_array
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using System; [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } [Tag("a"), Tag("b")] class Demo { } var attrs = (TagAttribute[])typeof(Demo).GetCustomAttributes(typeof(TagAttribute), false); foreach (var attr in attrs) Console.WriteLine(attr.Name);
