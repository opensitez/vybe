//! Init-only setters, object initializers with `init`, and `required` members (C# 11).

csharp_cases! {
    init_property_default_used_when_initializer_omits_it => {
        r#"class Config { public int Port { get; init; } = 8080; }
var c = new Config();
Console.WriteLine(c.Port);"#,
        ["8080"]
    };

    init_property_object_initializer_overrides_default => {
        r#"class Config { public int Port { get; init; } = 80; }
var c = new Config { Port = 443 };
Console.WriteLine(c.Port);"#,
        ["443"]
    };

    init_property_string_set_in_object_initializer => {
        r#"class User { public string Name { get; init; } = "guest"; }
var u = new User { Name = "Ada" };
Console.WriteLine(u.Name);"#,
        ["Ada"]
    };

    init_property_bool_set_in_object_initializer => {
        r#"class Flags { public bool Enabled { get; init; } }
var f = new Flags { Enabled = true };
Console.WriteLine(f.Enabled);"#,
        ["True"]
    };

    init_property_double_set_in_object_initializer => {
        r#"class Measure { public double Value { get; init; } }
var m = new Measure { Value = 3.5 };
Console.WriteLine(m.Value);"#,
        ["3.5"]
    };

    init_property_expression_computed_in_initializer => {
        r#"class Box { public int Size { get; init; } }
var b = new Box { Size = 10 + 5 };
Console.WriteLine(b.Size);"#,
        ["15"]
    };

    init_property_multiple_on_same_type => {
        r#"class Point { public int X { get; init; } public int Y { get; init; } }
var p = new Point { X = 2, Y = 7 };
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#,
        ["2", "7"]
    };

    init_property_on_struct_with_object_initializer => {
        r#"struct Pair { public int A { get; init; } public int B { get; init; } }
var p = new Pair { A = 4, B = 6 };
Console.WriteLine(p.A + p.B);"#,
        ["10"]
    };

    init_property_on_nominal_record_via_initializer => {
        r#"record Settings { public string Mode { get; init; } = "safe"; }
var s = new Settings { Mode = "fast" };
Console.WriteLine(s.Mode);"#,
        ["fast"]
    };

    init_property_on_positional_record_with_extra_init => {
        r#"record User(string Name) { public int Age { get; init; } = 0; }
var u = new User("Bob") { Age = 30 };
Console.WriteLine(u.Name); Console.WriteLine(u.Age);"#,
        ["Bob", "30"]
    };

    init_property_mixed_with_settable_property => {
        r#"class Item { public int Id { get; init; } public string Label { get; set; } = ""; }
var i = new Item { Id = 7 };
i.Label = "tool";
Console.WriteLine(i.Id); Console.WriteLine(i.Label);"#,
        ["7", "tool"]
    };

    init_property_and_public_field_in_same_initializer => {
        r#"class Form { public string Title { get; init; } public int Version; }
var f = new Form { Title = "main", Version = 2 };
Console.WriteLine(f.Title); Console.WriteLine(f.Version);"#,
        ["main", "2"]
    };

    init_property_inherited_and_set_in_derived_initializer => {
        r#"class Base { public string Tag { get; init; } = "base"; }
class Derived : Base { }
var d = new Derived { Tag = "child" };
Console.WriteLine(d.Tag);"#,
        ["child"]
    };

    init_property_nullable_int_set_in_initializer => {
        r#"class Maybe { public int? Count { get; init; } }
var m = new Maybe { Count = 5 };
Console.WriteLine(m.Count);"#,
        ["5"]
    };

    init_property_nullable_int_omitted_stays_null => {
        r#"class Maybe { public int? Count { get; init; } }
var m = new Maybe();
Console.WriteLine(m.Count.HasValue);"#,
        ["False"]
    };

    init_property_with_expression_default => {
        r#"class Scale { public int Factor { get; init; } = 2 * 3; }
var s = new Scale();
Console.WriteLine(s.Factor);"#,
        ["6"]
    };

    init_property_read_in_instance_method => {
        r#"class Config { public int Port { get; init; } = 80; public int DoublePort() => Port * 2; }
var c = new Config { Port = 11 };
Console.WriteLine(c.DoublePort());"#,
        ["22"]
    };

    init_property_on_nested_class => {
        r#"class Outer { public class Inner { public string Name { get; init; } } }
var i = new Outer.Inner { Name = "core" };
Console.WriteLine(i.Name);"#,
        ["core"]
    };

    record_struct_init_property_via_initializer => {
        r#"record struct Tag { public string Name { get; init; } = "none"; }
var t = new Tag { Name = "alpha" };
Console.WriteLine(t.Name);"#,
        ["alpha"]
    };

    init_property_enum_type_in_initializer => {
        r#"enum Level { Low, High }
class Job { public Level Tier { get; init; } = Level.Low; }
var j = new Job { Tier = Level.High };
Console.WriteLine(j.Tier);"#,
        ["High"]
    };

    init_property_char_type_in_initializer => {
        r#"class Token { public char Symbol { get; init; } = 'a'; }
var t = new Token { Symbol = 'z' };
Console.WriteLine(t.Symbol);"#,
        ["z"]
    };

    init_property_long_type_in_initializer => {
        r#"class Stats { public long Total { get; init; } }
var s = new Stats { Total = 10000000000L };
Console.WriteLine(s.Total);"#,
        ["10000000000"]
    };

    init_property_decimal_type_in_initializer => {
        r#"class Price { public decimal Amount { get; init; } }
var p = new Price { Amount = 19.99m };
Console.WriteLine(p.Amount);"#,
        ["19.99"]
    };

    init_property_date_time_in_initializer => {
        r#"class Event { public System.DateTime When { get; init; } }
var e = new Event { When = new System.DateTime(2024, 1, 15) };
Console.WriteLine(e.When.Year);"#,
        ["2024"]
    };

    init_property_array_reference_in_initializer => {
        r#"class Bundle { public int[] Items { get; init; } = new int[0]; }
var b = new Bundle { Items = new[] { 1, 2, 3 } };
Console.WriteLine(b.Items.Length);"#,
        ["3"]
    };

    init_property_list_reference_in_initializer => {
        r#"class Holder { public System.Collections.Generic.List<int> Values { get; init; } = new(); }
var h = new Holder { Values = new System.Collections.Generic.List<int> { 4, 5 } };
Console.WriteLine(h.Values.Count);"#,
        ["2"]
    };

    init_property_chained_parent_child_initializers => {
        r#"class Address { public string City { get; init; } }
class Person { public string Name { get; init; } public Address Home { get; init; } }
var p = new Person { Name = "Ann", Home = new Address { City = "Oslo" } };
Console.WriteLine(p.Home.City);"#,
        ["Oslo"]
    };

    with_expression_changes_init_property_on_record => {
        r#"record Config { public int Port { get; init; } = 80; }
var a = new Config();
var b = a with { Port = 9000 };
Console.WriteLine(a.Port); Console.WriteLine(b.Port);"#,
        ["80", "9000"]
    };

    init_property_two_instances_have_independent_values => {
        r#"class Slot { public int Id { get; init; } }
var a = new Slot { Id = 1 };
var b = new Slot { Id = 2 };
Console.WriteLine(a.Id); Console.WriteLine(b.Id);"#,
        ["1", "2"]
    };

    required_property_must_be_set_via_object_initializer => {
        r#"class Order { public required string Sku { get; set; } }
var o = new Order { Sku = "ABC" };
Console.WriteLine(o.Sku);"#,
        ["ABC"]
    };

    required_property_set_via_this_in_constructor => {
        r#"class Order { public required string Sku { get; set; } public Order(string sku) { Sku = sku; } }
var o = new Order("XYZ");
Console.WriteLine(o.Sku);"#,
        ["XYZ"]
    };

    required_property_on_record_set_in_initializer => {
        r#"record Person { public required string Name { get; init; } }
var p = new Person { Name = "Rex" };
Console.WriteLine(p.Name);"#,
        ["Rex"]
    };

    required_property_with_init_accessor => {
        r#"class Node { public required int Id { get; init; } }
var n = new Node { Id = 42 };
Console.WriteLine(n.Id);"#,
        ["42"]
    };

    required_field_set_in_object_initializer => {
        r#"class Packet { public required int Size; }
var p = new Packet { Size = 512 };
Console.WriteLine(p.Size);"#,
        ["512"]
    };

    required_field_set_in_constructor_body => {
        r#"class Packet { public required int Size; public Packet(int size) { Size = size; } }
var p = new Packet(256);
Console.WriteLine(p.Size);"#,
        ["256"]
    };

    required_field_on_struct_set_in_initializer => {
        r#"struct Block { public required int Length; }
var b = new Block { Length = 64 };
Console.WriteLine(b.Length);"#,
        ["64"]
    };

    multiple_required_members_on_same_class => {
        r#"class Pair { public required int Left; public required int Right; }
var p = new Pair { Left = 3, Right = 9 };
Console.WriteLine(p.Left + p.Right);"#,
        ["12"]
    };

    required_string_and_optional_init_property_together => {
        r#"class Profile { public required string User; public int Score { get; init; } = 0; }
var p = new Profile { User = "ada", Score = 100 };
Console.WriteLine(p.User); Console.WriteLine(p.Score);"#,
        ["ada", "100"]
    };

    required_enum_property_set_in_initializer => {
        r#"enum State { Off, On }
class Switch { public required State Mode { get; set; } }
var s = new Switch { Mode = State.On };
Console.WriteLine(s.Mode);"#,
        ["On"]
    };

    required_property_in_derived_class_initializer => {
        r#"class Base { public int Id; }
class Derived : Base { public required string Label { get; set; } }
var d = new Derived { Label = "child" };
Console.WriteLine(d.Label);"#,
        ["child"]
    };

    required_property_constructor_then_read => {
        r#"class Token { public required string Value { get; set; } public Token() { Value = "init"; } }
var t = new Token();
Console.WriteLine(t.Value);"#,
        ["init"]
    };

    init_property_default_preserved_across_two_instances => {
        r#"class Config { public int Retries { get; init; } = 3; }
var a = new Config();
var b = new Config { Retries = 1 };
Console.WriteLine(a.Retries); Console.WriteLine(b.Retries);"#,
        ["3", "1"]
    };

    init_property_used_in_equality_check => {
        r#"class Tag { public string Name { get; init; } = ""; }
var a = new Tag { Name = "x" };
var b = new Tag { Name = "x" };
Console.WriteLine(a.Name == b.Name);"#,
        ["True"]
    };

    init_property_string_empty_explicit_in_initializer => {
        r#"class Label { public string Text { get; init; } = "default"; }
var l = new Label { Text = "" };
Console.WriteLine(l.Text.Length);"#,
        ["0"]
    };

    init_property_byte_value_in_initializer => {
        r#"class ByteHolder { public byte Code { get; init; } }
var b = new ByteHolder { Code = 255 };
Console.WriteLine(b.Code);"#,
        ["255"]
    };

    init_property_short_value_in_initializer => {
        r#"class ShortHolder { public short Value { get; init; } }
var s = new ShortHolder { Value = 1000 };
Console.WriteLine(s.Value);"#,
        ["1000"]
    };

    init_property_float_value_in_initializer => {
        r#"class Sample { public float Rate { get; init; } }
var s = new Sample { Rate = 2.5f };
Console.WriteLine(s.Rate);"#,
        ["2.5"]
    };

    required_and_init_on_same_property_via_initializer => {
        r#"class Entity { public required int Id { get; init; } }
var e = new Entity { Id = 7 };
Console.WriteLine(e.Id);"#,
        ["7"]
    };

    init_property_object_initializer_partial_override_keeps_other_default => {
        r#"class Pair { public int A { get; init; } = 1; public int B { get; init; } = 2; }
var p = new Pair { B = 9 };
Console.WriteLine(p.A); Console.WriteLine(p.B);"#,
        ["1", "9"]
    };

    init_property_on_class_with_parameterless_constructor => {
        r#"class Widget { public Widget() { } public int Count { get; init; } = 0; }
var w = new Widget { Count = 5 };
Console.WriteLine(w.Count);"#,
        ["5"]
    };

    required_field_string_set_in_initializer => {
        r#"class Header { public required string Name; }
var h = new Header { Name = "Content-Type" };
Console.WriteLine(h.Name);"#,
        ["Content-Type"]
    };

    init_property_guid_value_in_initializer => {
        r#"class Ref { public System.Guid Id { get; init; } }
var id = new System.Guid("11111111-1111-1111-1111-111111111111");
var r = new Ref { Id = id };
Console.WriteLine(r.Id == id);"#,
        ["True"]
    };
}
