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
    nested_class_reads_outer_constant_value,
    r#"
class Outer {
    public const string Prefix = "outer";
    public class Inner {
        public string Read() { return Prefix + "/inner"; }
    }
}
Console.WriteLine(new Outer.Inner().Read());
"#,
    ["outer/inner"]
);

csharp_case!(
    nested_static_class_builds_formatter,
    r#"
class Report {
    public static class Formatter {
        public static string Line(string key, int value) { return key + ":" + value; }
    }
}
Console.WriteLine(Report.Formatter.Line("count", 3));
"#,
    ["count:3"]
);

csharp_case!(
    nested_enum_values_are_accessible_through_outer_type,
    r#"
class Job {
    public enum State { Pending, Running, Done }
}
Console.WriteLine(Job.State.Pending);
Console.WriteLine((int)Job.State.Done);
"#,
    ["Pending", "2"]
);

csharp_case!(
    nested_struct_can_be_created_from_outer_scope,
    r#"
class Geometry {
    public struct Point {
        public int X;
        public int Y;
    }
}
var point = new Geometry.Point { X = 3, Y = 4 };
Console.WriteLine(point.X + point.Y);
"#,
    ["7"]
);

csharp_case!(
    partial_class_combines_methods_from_two_parts,
    r#"
partial class Worker {
    public string First() { return "one"; }
}
partial class Worker {
    public string Second() { return "two"; }
}
var worker = new Worker();
Console.WriteLine(worker.First());
Console.WriteLine(worker.Second());
"#,
    ["one", "two"]
);

csharp_case!(
    partial_class_combines_field_and_method_declarations,
    r#"
partial class Config {
    string env = "prod";
}
partial class Config {
    public string Read() { return env; }
}
Console.WriteLine(new Config().Read());
"#,
    ["prod"]
);

csharp_case!(
    partial_class_combines_property_and_constructor_logic,
    r#"
partial class Build {
    public string Name { get; set; }
}
partial class Build {
    public Build(string name) { Name = name; }
}
Console.WriteLine(new Build("nightly").Name);
"#,
    ["nightly"]
);

csharp_case!(
    nested_class_inside_generic_outer_type_uses_type_argument,
    r#"
class Box<T> {
    public class Wrapper {
        public T Value { get; set; }
    }
}
var wrapper = new Box<int>.Wrapper { Value = 9 };
Console.WriteLine(wrapper.Value);
"#,
    ["9"]
);

csharp_case!(
    nested_interface_is_implemented_by_inner_class,
    r#"
class Device {
    public interface IPort { string Open(); }
    public class UsbPort : IPort {
        public string Open() { return "usb-open"; }
    }
}
Device.IPort port = new Device.UsbPort();
Console.WriteLine(port.Open());
"#,
    ["usb-open"]
);

csharp_case!(
    partial_class_methods_share_same_private_state,
    r#"
partial class Counter {
    int value;
}
partial class Counter {
    public void Bump() { value++; }
    public int Read() { return value; }
}
var counter = new Counter();
counter.Bump();
counter.Bump();
Console.WriteLine(counter.Read());
"#,
    ["2"]
);