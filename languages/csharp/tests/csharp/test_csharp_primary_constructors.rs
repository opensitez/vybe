//! Primary constructors on classes, structs, and records (C# 12).

csharp_cases! {
    primary_constructor_param_read_in_instance_method => {
        r#"class Counter(int start) {
    int current = start;
    public int Next() => ++current;
    public int Value => current;
}
var c = new Counter(10);
c.Next(); c.Next();
Console.WriteLine(c.Value);"#,
        ["12"]
    };

    primary_constructor_string_param_used_in_method => {
        r#"class Greeter(string prefix) {
    public string Greet(string name) => prefix + " " + name;
}
Console.WriteLine(new Greeter("Hello").Greet("World"));"#,
        ["Hello World"]
    };

    primary_constructor_struct_copies_params_to_fields => {
        r#"struct Point(int x, int y) {
    public int X = x;
    public int Y = y;
}
var p = new Point(3, 4);
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#,
        ["3", "4"]
    };

    primary_constructor_derived_passes_param_to_base => {
        r#"class Animal(string name) { public string Name => name; }
class Dog(string name, string breed) : Animal(name) { public string Breed => breed; }
var d = new Dog("Rex", "Lab");
Console.WriteLine(d.Name); Console.WriteLine(d.Breed);"#,
        ["Rex", "Lab"]
    };

    primary_constructor_param_used_in_property_getter => {
        r#"class Radius(int value) { public int Value => value; }
Console.WriteLine(new Radius(7).Value);"#,
        ["7"]
    };

    primary_constructor_two_params_summed_in_method => {
        r#"class Adder(int a, int b) { public int Sum() => a + b; }
Console.WriteLine(new Adder(3, 4).Sum());"#,
        ["7"]
    };

    primary_constructor_param_multiplied_in_method => {
        r#"class Scale(int factor) { public int Apply(int n) => n * factor; }
Console.WriteLine(new Scale(5).Apply(6));"#,
        ["30"]
    };

    primary_constructor_bool_param_controls_branch => {
        r#"class Gate(bool open) { public string State() => open ? "open" : "closed"; }
Console.WriteLine(new Gate(true).State());"#,
        ["open"]
    };

    primary_constructor_char_param_returned => {
        r#"class Symbol(char ch) { public char Value => ch; }
Console.WriteLine(new Symbol('Q').Value);"#,
        ["Q"]
    };

    primary_constructor_double_param_formatted => {
        r#"class Rate(double value) { public double Value => value; }
Console.WriteLine(new Rate(2.5).Value);"#,
        ["2.5"]
    };

    primary_constructor_decimal_param_stored => {
        r#"class Money(decimal amount) { public decimal Amount => amount; }
Console.WriteLine(new Money(9.99m).Amount);"#,
        ["9.99"]
    };

    primary_constructor_long_param_used => {
        r#"class Big(long n) { public long Value => n; }
Console.WriteLine(new Big(9000000000L).Value);"#,
        ["9000000000"]
    };

    primary_constructor_string_length_from_param => {
        r#"class Label(string text) { public int Length => text.Length; }
Console.WriteLine(new Label("abcd").Length);"#,
        ["4"]
    };

    primary_constructor_param_compared_to_literal => {
        r#"class Check(int n) { public bool IsTen() => n == 10; }
Console.WriteLine(new Check(10).IsTen());"#,
        ["True"]
    };

    primary_constructor_param_null_coalescing_default => {
        r#"class Maybe(string? text) { public string Safe() => text ?? "none"; }
Console.WriteLine(new Maybe(null).Safe());"#,
        ["none"]
    };

    primary_constructor_two_instances_are_independent => {
        r#"class Slot(int id) { public int Id => id; }
var a = new Slot(1);
var b = new Slot(2);
Console.WriteLine(a.Id); Console.WriteLine(b.Id);"#,
        ["1", "2"]
    };

    primary_constructor_record_positional_style => {
        r#"record Point(int X, int Y);
var p = new Point(1, 2);
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#,
        ["1", "2"]
    };

    primary_constructor_record_struct_value_type => {
        r#"record struct Coord(int X, int Y);
var c = new Coord(5, 6);
Console.WriteLine(c.X + c.Y);"#,
        ["11"]
    };

    primary_constructor_class_field_initialized_from_param => {
        r#"class Holder(int seed) { int value = seed; public int Read() => value; }
Console.WriteLine(new Holder(99).Read());"#,
        ["99"]
    };

    primary_constructor_param_used_in_string_interpolation => {
        r#"class Tag(string name) { public string Label() => $"tag:{name}"; }
Console.WriteLine(new Tag("core").Label());"#,
        ["tag:core"]
    };

    primary_constructor_param_passed_to_static_helper => {
        r#"class Worker(int id) {
    static string Format(int value) => "id=" + value;
    public string Show() => Format(id);
}
Console.WriteLine(new Worker(3).Show());"#,
        ["id=3"]
    };

    primary_constructor_generic_class_with_type_param => {
        r#"class Box<T>(T item) { public T Item => item; }
Console.WriteLine(new Box<int>(42).Item);"#,
        ["42"]
    };

    primary_constructor_generic_struct => {
        r#"struct Wrap<T>(T value) { public T Value => value; }
Console.WriteLine(new Wrap<string>("hi").Value);"#,
        ["hi"]
    };

    primary_constructor_param_in_conditional_expression => {
        r#"class Sign(int n) { public string Kind() => n >= 0 ? "pos" : "neg"; }
Console.WriteLine(new Sign(5).Kind());"#,
        ["pos"]
    };

    primary_constructor_negative_param_handled => {
        r#"class Sign(int n) { public int Abs() => n < 0 ? -n : n; }
Console.WriteLine(new Sign(-8).Abs());"#,
        ["8"]
    };

    primary_constructor_param_used_in_loop_count => {
        r#"class Repeat(int times) { public int Run() { int t = 0; for (int i = 0; i < times; i++) t++; return t; } }
Console.WriteLine(new Repeat(4).Run());"#,
        ["4"]
    };

    primary_constructor_array_param_length => {
        r#"class Pack(int[] data) { public int Count => data.Length; }
Console.WriteLine(new Pack(new[] { 1, 2, 3 }).Count);"#,
        ["3"]
    };

    primary_constructor_list_param_count => {
        r#"class Bag(System.Collections.Generic.List<int> items) { public int Count => items.Count; }
Console.WriteLine(new Bag(new System.Collections.Generic.List<int> { 1, 2 }).Count);"#,
        ["2"]
    };

    primary_constructor_param_in_switch_expression => {
        r#"class Mode(int code) { public string Name() => code switch { 1 => "a", 2 => "b", _ => "x" }; }
Console.WriteLine(new Mode(2).Name());"#,
        ["b"]
    };

    primary_constructor_three_params_chained_sum => {
        r#"class Triple(int a, int b, int c) { public int Total => a + b + c; }
Console.WriteLine(new Triple(1, 2, 3).Total);"#,
        ["6"]
    };

    primary_constructor_struct_copy_preserves_field_values => {
        r#"struct Pair(int a, int b) { public int A = a; public int B = b; }
var p = new Pair(2, 3);
var q = p;
Console.WriteLine(q.A + q.B);"#,
        ["5"]
    };

    primary_constructor_method_mutates_field_not_param => {
        r#"class Acc(int start) { int total = start; public void Add(int n) { total += n; } public int Value => total; }
var a = new Acc(1);
a.Add(4);
Console.WriteLine(a.Value);"#,
        ["5"]
    };

    primary_constructor_param_used_in_indexer => {
        r#"class Row(int size) {
    int[] cells = new int[size];
    public int this[int i] { get => cells[i]; set => cells[i] = value; }
}
var r = new Row(3);
r[1] = 9;
Console.WriteLine(r[1]);"#,
        ["9"]
    };

    primary_constructor_derived_adds_own_param => {
        r#"class Base(int x) { public int X => x; }
class Extra(int x, int y) : Base(x) { public int Y => y; }
Console.WriteLine(new Extra(2, 5).Y);"#,
        ["5"]
    };

    primary_constructor_interface_method_uses_param => {
        r#"interface IVal { int Get(); }
class Impl(int n) : IVal { public int Get() => n; }
IVal v = new Impl(12);
Console.WriteLine(v.Get());"#,
        ["12"]
    };

    primary_constructor_nested_type_on_outer => {
        r#"class Outer(int seed) {
    public class Inner { public int Value; }
    public Inner Make() => new Inner { Value = seed };
}
Console.WriteLine(new Outer(6).Make().Value);"#,
        ["6"]
    };

    primary_constructor_param_to_string_in_method => {
        r#"class Code(int n) { public string Text() => n.ToString(); }
Console.WriteLine(new Code(77).Text());"#,
        ["77"]
    };

    primary_constructor_enum_param_stored => {
        r#"enum Level { Low, High }
class Job(Level tier) { public Level Tier => tier; }
Console.WriteLine(new Job(Level.High).Tier);"#,
        ["High"]
    };

    primary_constructor_byte_param_value => {
        r#"class ByteBox(byte b) { public byte Value => b; }
Console.WriteLine(new ByteBox(200).Value);"#,
        ["200"]
    };

    primary_constructor_short_param_doubled => {
        r#"class ShortScale(short n) { public int Twice => n * 2; }
Console.WriteLine(new ShortScale(50).Twice);"#,
        ["100"]
    };

    primary_constructor_float_param_halved => {
        r#"class Half(float n) { public float Value => n / 2f; }
Console.WriteLine(new Half(10f).Value);"#,
        ["5"]
    };

    primary_constructor_string_param_uppercased => {
        r#"class Shout(string word) { public string Loud() => word.ToUpper(); }
Console.WriteLine(new Shout("go").Loud());"#,
        ["GO"]
    };

    primary_constructor_param_equality_between_instances => {
        r#"record Id(int Value);
var a = new Id(5);
var b = new Id(5);
Console.WriteLine(a == b);"#,
        ["True"]
    };

    primary_constructor_record_with_extra_method => {
        r#"record Point(int X, int Y) { public int Sum() => X + Y; }
Console.WriteLine(new Point(2, 3).Sum());"#,
        ["5"]
    };

    primary_constructor_class_with_static_factory => {
        r#"class Token(int id) {
    public int Id => id;
    public static Token Default() => new Token(0);
}
Console.WriteLine(Token.Default().Id);"#,
        ["0"]
    };

    primary_constructor_param_in_boolean_and_expression => {
        r#"class Flags(bool a, bool b) { public bool Both => a && b; }
Console.WriteLine(new Flags(true, true).Both);"#,
        ["True"]
    };

    primary_constructor_param_in_boolean_or_expression => {
        r#"class Flags(bool a, bool b) { public bool Any => a || b; }
Console.WriteLine(new Flags(false, true).Any);"#,
        ["True"]
    };

    primary_constructor_string_is_null_or_empty_check => {
        r#"class Name(string? text) { public bool Missing => string.IsNullOrEmpty(text); }
Console.WriteLine(new Name("").Missing);"#,
        ["True"]
    };

    primary_constructor_param_modulo_operation => {
        r#"class Mod(int n) { public int Rem(int d) => n % d; }
Console.WriteLine(new Mod(10).Rem(3));"#,
        ["1"]
    };

    primary_constructor_param_bitwise_and => {
        r#"class Mask(int n) { public int And(int m) => n & m; }
Console.WriteLine(new Mask(12).And(10));"#,
        ["8"]
    };

    primary_constructor_param_shift_left => {
        r#"class Shift(int n) { public int Left(int bits) => n << bits; }
Console.WriteLine(new Shift(3).Left(2));"#,
        ["12"]
    };
}
