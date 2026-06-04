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
    qualified_type_name_instantiates_class_from_namespace,
    r#"namespace Demo { public class Box { public string Name => "demo"; } } var box = new Demo.Box(); Console.WriteLine(box.Name);"#,
    ["demo"]
);
csharp_case!(
    nested_namespace_type_is_reachable_by_full_name,
    r#"namespace Outer.Inner { public class Worker { public string Run() { return "ok"; } } } Console.WriteLine(new Outer.Inner.Worker().Run());"#,
    ["ok"]
);
csharp_case!(
    using_alias_can_shorten_fully_qualified_type_name,
    r#"using Thing = Demo.Tools.Box; namespace Demo.Tools { public class Box { public int Value = 7; } } Console.WriteLine(new Thing().Value);"#,
    ["7"]
);
csharp_case!(
    using_alias_can_target_generic_type,
    r#"using TextList = System.Collections.Generic.List<string>; var list = new TextList { "a", "b" }; Console.WriteLine(list.Count);"#,
    ["2"]
);
csharp_case!(
    using_static_imports_math_members_for_direct_calls,
    r#"using static System.Math; Console.WriteLine(Max(3, 9));"#,
    ["9"]
);
csharp_case!(
    using_static_imports_console_write_line_symbol,
    r#"using static System.Console; WriteLine("hello");"#,
    ["hello"]
);
csharp_case!(
    distinct_namespaces_can_define_same_type_name,
    r#"namespace Left { public class Item { public string Name => "L"; } } namespace Right { public class Item { public string Name => "R"; } } Console.WriteLine(new Left.Item().Name); Console.WriteLine(new Right.Item().Name);"#,
    ["L", "R"]
);
csharp_case!(
    namespace_scoped_enum_is_resolved_by_qualified_name,
    r#"namespace Demo { public enum State { Ready } } Console.WriteLine(Demo.State.Ready);"#,
    ["Ready"]
);
csharp_case!(
    namespace_scoped_interface_is_implemented_by_qualified_type,
    r#"namespace Demo { public interface IRun { string Run(); } public class Worker : IRun { public string Run() { return "done"; } } } Demo.IRun worker = new Demo.Worker(); Console.WriteLine(worker.Run());"#,
    ["done"]
);
csharp_case!(
    using_alias_for_namespace_selects_nested_type,
    r#"using Core = Demo.Core; namespace Demo.Core { public class Item { public string Name => "core"; } } Console.WriteLine(new Core.Item().Name);"#,
    ["core"]
);
csharp_case!(
    global_system_namespace_type_is_available_inside_custom_namespace,
    r#"namespace Demo { public class Worker { public string Read() { return global::System.String.Join(",", new[] { "a", "b" }); } } } Console.WriteLine(new Demo.Worker().Read());"#,
    ["a,b"]
);
csharp_case!(
    namespace_can_contain_nested_struct_type,
    r#"namespace Demo { public struct Point { public int X; public int Y; } } var point = new Demo.Point { X = 2, Y = 5 }; Console.WriteLine(point.X + point.Y);"#,
    ["7"]
);
csharp_case!(
    namespace_can_contain_static_helper_class,
    r#"namespace Demo.Tools { public static class MathEx { public static int Double(int value) { return value * 2; } } } Console.WriteLine(Demo.Tools.MathEx.Double(6));"#,
    ["12"]
);
csharp_case!(
    alias_can_reference_nested_class_type,
    r#"using InnerType = Demo.Outer.Inner; namespace Demo { public class Outer { public class Inner { public string Name => "inner"; } } } Console.WriteLine(new InnerType().Name);"#,
    ["inner"]
);
csharp_case!(
    using_directive_imports_custom_namespace_for_unqualified_access,
    r#"using Demo.Tools; namespace Demo.Tools { public class Worker { public string Name => "tool"; } } Console.WriteLine(new Worker().Name);"#,
    ["tool"]
);
csharp_case!(
    namespace_and_class_can_share_root_name_without_ambiguity,
    r#"namespace Demo.Sub { public class Demo { public int Value = 5; } } Console.WriteLine(new Demo.Sub.Demo().Value);"#,
    ["5"]
);
csharp_case!(
    qualified_reference_accesses_nested_enum_member,
    r#"namespace Demo { public class Job { public enum State { Idle, Done } } } Console.WriteLine(Demo.Job.State.Done);"#,
    ["Done"]
);
csharp_case!(
    namespace_scoped_delegate_can_be_invoked_through_qualified_name,
    r#"namespace Demo { public delegate string Reader(); } Demo.Reader reader = () => "text"; Console.WriteLine(reader());"#,
    ["text"]
);
csharp_case!(
    multiple_using_directives_can_import_separate_namespaces,
    r#"using Demo.Left; using Demo.Right; namespace Demo.Left { public class A { public string Name => "A"; } } namespace Demo.Right { public class B { public string Name => "B"; } } Console.WriteLine(new A().Name + new B().Name);"#,
    ["AB"]
);
csharp_case!(
    namespace_scoped_attribute_can_be_applied_by_short_name_after_using,
    r#"using Demo; namespace Demo { public class TagAttribute : System.Attribute { public string Name; public TagAttribute(string name) { Name = name; } } } [Tag("x")] class Item { } var attr = (Demo.TagAttribute)System.Attribute.GetCustomAttribute(typeof(Item), typeof(Demo.TagAttribute)); Console.WriteLine(attr.Name);"#,
    ["x"]
);
