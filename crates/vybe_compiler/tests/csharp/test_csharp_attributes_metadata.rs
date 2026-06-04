use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    custom_attribute_on_class_can_be_read_via_reflection,
    r#"using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public LabelAttribute(string name) { Name = name; } } [Label("service")] class Worker { } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(Worker), typeof(LabelAttribute)); Console.WriteLine(attr.Name);"#,
    ["service"]
);
csharp_case!(
    custom_attribute_on_method_can_be_read_via_reflection,
    r#"using System; [AttributeUsage(AttributeTargets.Method)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } class Worker { [Tag("run")] public void Execute() { } } var method = typeof(Worker).GetMethod("Execute"); var attr = (TagAttribute)Attribute.GetCustomAttribute(method, typeof(TagAttribute)); Console.WriteLine(attr.Name);"#,
    ["run"]
);
csharp_case!(
    custom_attribute_on_property_is_visible_through_property_info,
    r#"using System; [AttributeUsage(AttributeTargets.Property)] class HintAttribute : Attribute { public string Text { get; } public HintAttribute(string text) { Text = text; } } class Settings { [Hint("port")] public int Port { get; set; } } var property = typeof(Settings).GetProperty("Port"); var attr = (HintAttribute)Attribute.GetCustomAttribute(property, typeof(HintAttribute)); Console.WriteLine(attr.Text);"#,
    ["port"]
);
csharp_case!(
    custom_attribute_on_field_is_visible_through_field_info,
    r#"using System; [AttributeUsage(AttributeTargets.Field)] class MarkerAttribute : Attribute { public int Code { get; } public MarkerAttribute(int code) { Code = code; } } class Flags { [Marker(7)] public int Value; } var field = typeof(Flags).GetField("Value"); var attr = (MarkerAttribute)Attribute.GetCustomAttribute(field, typeof(MarkerAttribute)); Console.WriteLine(attr.Code);"#,
    ["7"]
);
csharp_case!(
    custom_attribute_on_parameter_is_visible_through_parameter_info,
    r#"using System; [AttributeUsage(AttributeTargets.Parameter)] class UnitAttribute : Attribute { public string Name { get; } public UnitAttribute(string name) { Name = name; } } class MathOps { public void Scale([Unit("px")] int value) { } } var parameter = typeof(MathOps).GetMethod("Scale").GetParameters()[0]; var attr = (UnitAttribute)Attribute.GetCustomAttribute(parameter, typeof(UnitAttribute)); Console.WriteLine(attr.Name);"#,
    ["px"]
);
csharp_case!(
    attribute_named_argument_sets_property_value,
    r#"using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public int Priority { get; set; } public LabelAttribute(string name) { Name = name; } } [Label("job", Priority = 3)] class TaskItem { } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(TaskItem), typeof(LabelAttribute)); Console.WriteLine(attr.Name); Console.WriteLine(attr.Priority);"#,
    ["job", "3"]
);
csharp_case!(
    attribute_allow_multiple_returns_two_instances,
    r#"using System; [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } [Tag("api"), Tag("internal")] class Endpoint { } var attrs = typeof(Endpoint).GetCustomAttributes(typeof(TagAttribute), false); Console.WriteLine(attrs.Length);"#,
    ["2"]
);
csharp_case!(
    attribute_inheritance_flows_to_derived_type_when_enabled,
    r#"using System; [AttributeUsage(AttributeTargets.Class, Inherited = true)] class RoleAttribute : Attribute { public string Name { get; } public RoleAttribute(string name) { Name = name; } } [Role("base")] class BaseController { } class DerivedController : BaseController { } var attr = (RoleAttribute)Attribute.GetCustomAttribute(typeof(DerivedController), typeof(RoleAttribute)); Console.WriteLine(attr.Name);"#,
    ["base"]
);
csharp_case!(
    obsolete_attribute_still_allows_method_invocation,
    r#"using System; class Service { [Obsolete("legacy")] public string Run() { return "ok"; } } Console.WriteLine(new Service().Run());"#,
    ["ok"]
);
csharp_case!(
    flags_attribute_marks_enum_but_bitwise_result_still_formats_value,
    r#"using System; [Flags] enum Permission { Read = 1, Write = 2, Execute = 4 } var permission = Permission.Read | Permission.Write; Console.WriteLine(permission);"#,
    ["Read, Write"]
);
csharp_case!(
    serializable_attribute_is_detectable_on_type,
    r#"using System; [Serializable] class Packet { } Console.WriteLine(Attribute.IsDefined(typeof(Packet), typeof(SerializableAttribute)));"#,
    ["True"]
);
csharp_case!(
    clscompliant_attribute_is_detectable_on_type,
    r#"using System; [CLSCompliant(true)] class PublicApi { } Console.WriteLine(Attribute.IsDefined(typeof(PublicApi), typeof(CLSCompliantAttribute)));"#,
    ["True"]
);
csharp_case!(
    attribute_constructor_can_capture_integer_argument,
    r#"using System; [AttributeUsage(AttributeTargets.Class)] class CodeAttribute : Attribute { public int Value { get; } public CodeAttribute(int value) { Value = value; } } [Code(42)] class Job { } var attr = (CodeAttribute)Attribute.GetCustomAttribute(typeof(Job), typeof(CodeAttribute)); Console.WriteLine(attr.Value);"#,
    ["42"]
);
csharp_case!(
    attribute_on_nested_class_is_readable_via_type_handle,
    r#"using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public LabelAttribute(string name) { Name = name; } } class Outer { [Label("inner")] public class Inner { } } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(Outer.Inner), typeof(LabelAttribute)); Console.WriteLine(attr.Name);"#,
    ["inner"]
);
csharp_case!(
    attribute_on_interface_is_readable_via_reflection,
    r#"using System; [AttributeUsage(AttributeTargets.Interface)] class ContractAttribute : Attribute { public string Name { get; } public ContractAttribute(string name) { Name = name; } } [Contract("service")] interface IService { } var attr = (ContractAttribute)Attribute.GetCustomAttribute(typeof(IService), typeof(ContractAttribute)); Console.WriteLine(attr.Name);"#,
    ["service"]
);
csharp_case!(
    attribute_on_struct_is_readable_via_reflection,
    r#"using System; [AttributeUsage(AttributeTargets.Struct)] class ShapeAttribute : Attribute { public string Name { get; } public ShapeAttribute(string name) { Name = name; } } [Shape("point")] struct Point { } var attr = (ShapeAttribute)Attribute.GetCustomAttribute(typeof(Point), typeof(ShapeAttribute)); Console.WriteLine(attr.Name);"#,
    ["point"]
);
csharp_case!(
    attribute_on_enum_is_readable_via_reflection,
    r#"using System; [AttributeUsage(AttributeTargets.Enum)] class GroupAttribute : Attribute { public string Name { get; } public GroupAttribute(string name) { Name = name; } } [Group("status")] enum State { Idle } var attr = (GroupAttribute)Attribute.GetCustomAttribute(typeof(State), typeof(GroupAttribute)); Console.WriteLine(attr.Name);"#,
    ["status"]
);
csharp_case!(
    attribute_get_custom_attributes_can_return_strongly_typed_array,
    r#"using System; [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } [Tag("a"), Tag("b")] class Demo { } var attrs = (TagAttribute[])typeof(Demo).GetCustomAttributes(typeof(TagAttribute), false); foreach (var attr in attrs) Console.WriteLine(attr.Name);"#,
    ["a", "b"]
);
csharp_case!(
    attribute_is_defined_reports_false_when_missing,
    r#"using System; class Plain { } Console.WriteLine(Attribute.IsDefined(typeof(Plain), typeof(ObsoleteAttribute)));"#,
    ["False"]
);
csharp_case!(
    attribute_can_be_read_from_base_method_definition,
    r#"using System; [AttributeUsage(AttributeTargets.Method)] class InfoAttribute : Attribute { public string Name { get; } public InfoAttribute(string name) { Name = name; } } class Base { [Info("root")] public virtual void Run() { } } class Derived : Base { public override void Run() { } } var method = typeof(Base).GetMethod("Run"); var attr = (InfoAttribute)Attribute.GetCustomAttribute(method, typeof(InfoAttribute)); Console.WriteLine(attr.Name);"#,
    ["root"]
);
