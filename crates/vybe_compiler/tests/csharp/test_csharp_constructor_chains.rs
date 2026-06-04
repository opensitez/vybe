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
    base_constructor_initializes_inherited_field,
    r#"class Base { protected string name; public Base(string name) { this.name = name; } public string Name() { return name; } } class Child : Base { public Child() : base("root") { } } Console.WriteLine(new Child().Name());"#,
    ["root"]
);
csharp_case!(
    this_constructor_chain_reuses_primary_overload,
    r#"class Box { int value; public Box() : this(9) { } public Box(int value) { this.value = value; } public int Read() { return value; } } Console.WriteLine(new Box().Read());"#,
    ["9"]
);
csharp_case!(
    constructor_can_set_readonly_field_from_parameter,
    r#"class Box { readonly int value; public Box(int value) { this.value = value; } public int Read() { return value; } } Console.WriteLine(new Box(7).Read());"#,
    ["7"]
);
csharp_case!(
    field_initializer_runs_before_constructor_body_reads_value,
    r#"class Box { string name = "init"; public Box() { Console.WriteLine(name); } } new Box();"#,
    ["init"]
);
csharp_case!(
    constructor_overload_can_append_suffix_after_chain,
    r#"class Box { string name; public Box(string name) { this.name = name; } public Box(string name, string suffix) : this(name) { this.name += suffix; } public string Read() { return name; } } Console.WriteLine(new Box("a", "b").Read());"#,
    ["ab"]
);
csharp_case!(
    private_constructor_can_be_reached_through_factory_method,
    r#"class Box { string name; private Box(string name) { this.name = name; } public static Box Create() { return new Box("made"); } public string Read() { return name; } } Console.WriteLine(Box.Create().Read());"#,
    ["made"]
);
csharp_case!(
    base_and_derived_constructors_run_in_order,
    r#"class Base { public Base() { Console.WriteLine("base"); } } class Child : Base { public Child() { Console.WriteLine("child"); } } new Child();"#,
    ["base", "child"]
);
csharp_case!(
    constructor_can_accept_optional_parameter_value,
    r#"class Box { int value; public Box(int value = 6) { this.value = value; } public int Read() { return value; } } Console.WriteLine(new Box().Read());"#,
    ["6"]
);
csharp_case!(
    constructor_can_initialize_auto_property,
    r#"class Box { public string Name { get; } public Box(string name) { Name = name; } } Console.WriteLine(new Box("pkg").Name);"#,
    ["pkg"]
);
csharp_case!(
    struct_constructor_can_delegate_to_field_assignment,
    r#"struct Pair { public int Left; public int Right; public Pair(int left, int right) { Left = left; Right = right; } } var pair = new Pair(2, 8); Console.WriteLine(pair.Left + pair.Right);"#,
    ["10"]
);
csharp_case!(
    derived_constructor_can_pass_argument_from_expression_to_base,
    r#"class Base { int value; public Base(int value) { this.value = value; } public int Read() { return value; } } class Child : Base { public Child(int value) : base(value + 1) { } } Console.WriteLine(new Child(4).Read());"#,
    ["5"]
);
csharp_case!(
    constructor_can_initialize_collection_field,
    r#"using System.Collections.Generic; class Box { List<int> values; public Box() { values = new List<int> { 1, 2, 3 }; } public int Count() { return values.Count; } } Console.WriteLine(new Box().Count());"#,
    ["3"]
);
csharp_case!(
    nested_class_constructor_can_capture_argument,
    r#"class Outer { public class Inner { string name; public Inner(string name) { this.name = name; } public string Read() { return name; } } } Console.WriteLine(new Outer.Inner("inner").Read());"#,
    ["inner"]
);
csharp_case!(
    constructor_can_call_instance_method_after_assignment,
    r#"class Box { int value; public Box(int value) { this.value = value; Console.WriteLine(Read()); } public int Read() { return value; } } new Box(8);"#,
    ["8"]
);
csharp_case!(
    object_initializer_runs_after_constructor_default_assignment,
    r#"class Box { public string Name { get; set; } public Box() { Name = "init"; } } var box = new Box { Name = "set" }; Console.WriteLine(box.Name);"#,
    ["set"]
);
csharp_case!(
    generic_class_constructor_can_store_type_specific_value,
    r#"class Box<T> { T value; public Box(T value) { this.value = value; } public T Read() { return value; } } Console.WriteLine(new Box<string>("text").Read());"#,
    ["text"]
);
csharp_case!(
    base_constructor_and_override_dispatch_can_coexist_after_construction,
    r#"class Base { protected string prefix; public Base(string prefix) { this.prefix = prefix; } public virtual string Read() { return prefix; } } class Child : Base { public Child() : base("x") { } public override string Read() { return prefix + "y"; } } Console.WriteLine(new Child().Read());"#,
    ["xy"]
);
csharp_case!(
    static_constructor_and_instance_constructor_both_run_for_first_instance,
    r#"class Box { static Box() { Console.WriteLine("static"); } public Box() { Console.WriteLine("instance"); } } new Box();"#,
    ["static", "instance"]
);
csharp_case!(
    constructor_chain_can_set_multiple_fields_from_single_input,
    r#"class Box { string left; string right; public Box(string value) : this(value, value.ToUpper()) { } public Box(string left, string right) { this.left = left; this.right = right; } public string Read() { return left + ":" + right; } } Console.WriteLine(new Box("a").Read());"#,
    ["a:A"]
);
csharp_case!(
    record_primary_constructor_members_are_available_immediately,
    r#"record User(string Name, int Age); var user = new User("Ada", 20); Console.WriteLine(user.Name); Console.WriteLine(user.Age);"#,
    ["Ada", "20"]
);
