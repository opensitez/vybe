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
    explicit_interface_method_is_called_via_interface_reference,
    r#"
interface IGreeter { string Speak(); }
class Person : IGreeter {
    string IGreeter.Speak() { return "hello"; }
}
IGreeter greeter = new Person();
Console.WriteLine(greeter.Speak());
"#,
    ["hello"]
);

csharp_case!(
    explicit_interface_property_is_visible_through_interface,
    r#"
interface IValueHolder { int Value { get; } }
class Counter : IValueHolder {
    int IValueHolder.Value { get { return 12; } }
}
IValueHolder holder = new Counter();
Console.WriteLine(holder.Value);
"#,
    ["12"]
);

csharp_case!(
    explicit_implementations_disambiguate_same_method_name,
    r#"
interface IText { string Format(); }
interface IJson { string Format(); }
class Payload : IText, IJson {
    string IText.Format() { return "text"; }
    string IJson.Format() { return "json"; }
}
var payload = new Payload();
Console.WriteLine(((IText)payload).Format());
Console.WriteLine(((IJson)payload).Format());
"#,
    ["text", "json"]
);

csharp_case!(
    explicit_interface_method_coexists_with_public_method,
    r#"
interface IFormatter { string Format(); }
class Report : IFormatter {
    public string Format() { return "public"; }
    string IFormatter.Format() { return "explicit"; }
}
var report = new Report();
Console.WriteLine(report.Format());
Console.WriteLine(((IFormatter)report).Format());
"#,
    ["public", "explicit"]
);

csharp_case!(
    explicit_interface_method_is_invoked_after_cast_from_object,
    r#"
interface IRunner { string Run(); }
class TaskRunner : IRunner {
    string IRunner.Run() { return "done"; }
}
object item = new TaskRunner();
Console.WriteLine(((IRunner)item).Run());
"#,
    ["done"]
);

csharp_case!(
    explicit_interface_method_on_generic_interface_returns_value,
    r#"
interface IBox<T> { T Unwrap(); }
class NumberBox : IBox<int> {
    int IBox<int>.Unwrap() { return 42; }
}
IBox<int> box = new NumberBox();
Console.WriteLine(box.Unwrap());
"#,
    ["42"]
);

csharp_case!(
    explicit_interface_property_and_method_share_private_state,
    r#"
interface IStatus {
    string Name { get; }
    string Read();
}
class Job : IStatus {
    string name = "queued";
    string IStatus.Name { get { return name; } }
    string IStatus.Read() { return name + "!"; }
}
IStatus status = new Job();
Console.WriteLine(status.Name);
Console.WriteLine(status.Read());
"#,
    ["queued", "queued!"]
);

csharp_case!(
    explicit_interface_method_works_with_base_class_inheritance,
    r#"
interface ILabel { string Label(); }
class BaseItem {
    protected string prefix = "base";
}
class TaggedItem : BaseItem, ILabel {
    string ILabel.Label() { return prefix + "/tag"; }
}
Console.WriteLine(((ILabel)new TaggedItem()).Label());
"#,
    ["base/tag"]
);

csharp_case!(
    explicit_interface_indexer_reads_values_through_interface,
    r#"
interface IReadIndex { string this[int index] { get; } }
class Words : IReadIndex {
    string[] values = new[] { "alpha", "beta" };
    string IReadIndex.this[int index] { get { return values[index]; } }
}
IReadIndex words = new Words();
Console.WriteLine(words[0]);
Console.WriteLine(words[1]);
"#,
    ["alpha", "beta"]
);

csharp_case!(
    explicit_interface_can_be_accessed_after_multiple_casts,
    r#"
interface ICode { string Value(); }
class Ticket : ICode {
    string ICode.Value() { return "T-9"; }
}
var ticket = new Ticket();
object boxed = ticket;
Console.WriteLine(((ICode)boxed).Value());
"#,
    ["T-9"]
);
