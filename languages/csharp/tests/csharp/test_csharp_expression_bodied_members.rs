//! Expression-bodied members: methods, properties, operators, and indexers on classes and structs.

csharp_cases! {
    expr_method_class_returns_doubled_int => {
        r#"class Calc { public int Double(int n) => n * 2; }
Console.WriteLine(new Calc().Double(5));"#,
        ["10"]
    };

    expr_method_class_void_writes_argument => {
        r#"class Echo { public void Say(string msg) => Console.WriteLine(msg); }
new Echo().Say("hi");"#,
        ["hi"]
    };

    expr_method_class_static_clamps_value => {
        r#"static class ClampUtil { public static int Clamp(int v, int lo, int hi) => v < lo ? lo : v > hi ? hi : v; }
Console.WriteLine(ClampUtil.Clamp(15, 0, 10));"#,
        ["10"]
    };

    expr_method_class_two_params_sums => {
        r#"class Adder { public int Sum(int a, int b) => a + b; }
Console.WriteLine(new Adder().Sum(3, 4));"#,
        ["7"]
    };

    expr_method_class_three_params_product => {
        r#"class Mul3 { public int Prod(int a, int b, int c) => a * b * c; }
Console.WriteLine(new Mul3().Prod(2, 3, 4));"#,
        ["24"]
    };

    expr_method_class_returns_string_concat => {
        r#"class Joiner { public string Merge(string a, string b) => a + b; }
Console.WriteLine(new Joiner().Merge("ab", "cd"));"#,
        ["abcd"]
    };

    expr_method_class_returns_bool_comparison => {
        r#"class Check { public bool IsZero(int n) => n == 0; }
Console.WriteLine(new Check().IsZero(0)); Console.WriteLine(new Check().IsZero(1));"#,
        ["True", "False"]
    };

    expr_method_class_returns_char_from_index => {
        r#"class Pick { public char At(string s, int i) => s[i]; }
Console.WriteLine(new Pick().At("cat", 1));"#,
        ["a"]
    };

    expr_method_class_returns_double_ratio => {
        r#"class Ratio { public double Half(double x) => x / 2.0; }
Console.WriteLine(new Ratio().Half(5.0));"#,
        ["2.5"]
    };

    expr_method_struct_instance_on_stack => {
        r#"struct Counter { public int n; public int Next() => ++n; }
var c = new Counter();
Console.WriteLine(c.Next()); Console.WriteLine(c.Next());"#,
        ["1", "2"]
    };

    expr_method_struct_static_factory => {
        r#"struct Point { public int X, Y; public static Point Origin() => new Point { X = 0, Y = 0 }; }
var p = Point.Origin();
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#,
        ["0", "0"]
    };

    expr_method_nested_class_delegates_to_outer => {
        r#"class Outer { public int Base => 10; public class Inner { Outer o; public Inner(Outer owner) { o = owner; } public int Boost() => o.Base + 5; } }
Console.WriteLine(new Outer.Inner(new Outer()).Boost());"#,
        ["15"]
    };

    expr_method_override_expression_body => {
        r#"class Base { public virtual int Id() => 1; }
class Derived : Base { public override int Id() => 2; }
Console.WriteLine(new Derived().Id());"#,
        ["2"]
    };

    expr_method_class_uses_field_in_body => {
        r#"class Scale { public int factor = 3; public int Apply(int n) => n * factor; }
Console.WriteLine(new Scale().Apply(4));"#,
        ["12"]
    };

    expr_method_class_null_coalescing_param => {
        r#"class Safe { public string OrEmpty(string? s) => s ?? ""; }
Console.WriteLine(new Safe().OrEmpty(null)); Console.WriteLine(new Safe().OrEmpty("x"));"#,
        ["", "x"]
    };

    expr_property_class_get_only_from_field => {
        r#"class Circle { public double R = 2.0; public double Area => System.Math.PI * R * R; }
Console.WriteLine(System.Math.Round(new Circle().Area, 2));"#,
        ["12.57"]
    };

    expr_property_class_get_set_both_expression => {
        r#"class Box { int _v; public int Value { get => _v; set => _v = value; } }
var b = new Box(); b.Value = 9; Console.WriteLine(b.Value);"#,
        ["9"]
    };

    expr_property_class_static_readonly_computed => {
        r#"static class Consts { public static int Ten => 10; }
Console.WriteLine(Consts.Ten);"#,
        ["10"]
    };

    expr_property_class_string_length_computed => {
        r#"class Label { public string Text = "hello"; public int Len => Text.Length; }
Console.WriteLine(new Label().Len);"#,
        ["5"]
    };

    expr_property_class_bool_is_empty => {
        r#"class Bag { public string? Data; public bool IsEmpty => Data == null || Data.Length == 0; }
Console.WriteLine(new Bag { Data = "" }.IsEmpty); Console.WriteLine(new Bag { Data = "x" }.IsEmpty);"#,
        ["True", "False"]
    };

    expr_property_struct_get_only => {
        r#"struct Pair { public int A, B; public int Sum => A + B; }
var p = new Pair { A = 2, B = 5 }; Console.WriteLine(p.Sum);"#,
        ["7"]
    };

    expr_property_struct_get_set => {
        r#"struct Slot { int _n; public int N { get => _n; set => _n = value; } }
var s = new Slot(); s.N = 7; Console.WriteLine(s.N);"#,
        ["7"]
    };

    expr_property_class_chained_computed => {
        r#"class Chain { public int Base = 2; public int Double => Base * 2; public int Quadruple => Double * 2; }
Console.WriteLine(new Chain().Quadruple);"#,
        ["8"]
    };

    expr_property_class_char_upper_from_field => {
        r#"class Token { public char ch = 'a'; public char Upper => char.ToUpper(ch); }
Console.WriteLine(new Token().Upper);"#,
        ["A"]
    };

    expr_property_class_index_from_one_based => {
        r#"class Row { public int Index = 0; public int Display => Index + 1; }
Console.WriteLine(new Row { Index = 4 }.Display);"#,
        ["5"]
    };

    expr_property_class_percent_full => {
        r#"class Tank { public int level = 75; public int capacity = 100; public int Percent => level * 100 / capacity; }
Console.WriteLine(new Tank().Percent);"#,
        ["75"]
    };

    expr_property_class_setter_expression_updates_field => {
        r#"class Logger { public string last = ""; public string Last { get => last; set => last = value; } }
var l = new Logger(); l.Last = "ok"; Console.WriteLine(l.Last);"#,
        ["ok"]
    };

    expr_operator_class_addition => {
        r#"class Num { public int V; public static Num operator +(Num a, Num b) => new Num { V = a.V + b.V }; }
Console.WriteLine((new Num { V = 3 } + new Num { V = 4 }).V);"#,
        ["7"]
    };

    expr_operator_class_subtraction => {
        r#"class Num { public int V; public static Num operator -(Num a, Num b) => new Num { V = a.V - b.V }; }
Console.WriteLine((new Num { V = 10 } - new Num { V = 4 }).V);"#,
        ["6"]
    };

    expr_operator_class_multiplication => {
        r#"class Num { public int V; public static Num operator *(Num a, Num b) => new Num { V = a.V * b.V }; }
Console.WriteLine((new Num { V = 3 } * new Num { V = 5 }).V);"#,
        ["15"]
    };

    expr_operator_class_equality => {
        r#"class Tag { public string Name; public static bool operator ==(Tag a, Tag b) => a.Name == b.Name; public static bool operator !=(Tag a, Tag b) => !(a == b); }
Console.WriteLine(new Tag { Name = "x" } == new Tag { Name = "x" }); Console.WriteLine(new Tag { Name = "a" } != new Tag { Name = "b" });"#,
        ["True", "True"]
    };

    expr_operator_class_less_than => {
        r#"class Score { public int V; public static bool operator <(Score a, Score b) => a.V < b.V; public static bool operator >(Score a, Score b) => a.V > b.V; }
Console.WriteLine(new Score { V = 1 } < new Score { V = 2 }); Console.WriteLine(new Score { V = 5 } > new Score { V = 3 });"#,
        ["True", "True"]
    };

    expr_operator_struct_addition => {
        r#"struct Vec2 { public int X, Y; public static Vec2 operator +(Vec2 a, Vec2 b) => new Vec2 { X = a.X + b.X, Y = a.Y + b.Y }; }
var v = new Vec2 { X = 1, Y = 2 } + new Vec2 { X = 3, Y = 4 };
Console.WriteLine(v.X); Console.WriteLine(v.Y);"#,
        ["4", "6"]
    };

    expr_operator_struct_unary_minus => {
        r#"struct Signed { public int V; public static Signed operator -(Signed s) => new Signed { V = -s.V }; }
Console.WriteLine((-new Signed { V = 7 }).V);"#,
        ["-7"]
    };

    expr_operator_struct_unary_plus => {
        r#"struct Signed { public int V; public static Signed operator +(Signed s) => new Signed { V = +s.V }; }
Console.WriteLine((+new Signed { V = 7 }).V);"#,
        ["7"]
    };

    expr_operator_implicit_conversion_to_int => {
        r#"struct Wrap { public int V; public static implicit operator int(Wrap w) => w.V; }
Wrap w = new Wrap { V = 42 }; int n = w; Console.WriteLine(n);"#,
        ["42"]
    };

    expr_operator_explicit_conversion_from_int => {
        r#"struct Wrap { public int V; public static explicit operator Wrap(int n) => new Wrap { V = n }; }
Wrap w = (Wrap)9; Console.WriteLine(w.V);"#,
        ["9"]
    };

    expr_operator_true_false_for_custom_type => {
        r#"struct Flag { public bool On; public static bool operator true(Flag f) => f.On; public static bool operator false(Flag f) => !f.On; }
Flag f = new Flag { On = true }; if (f) Console.WriteLine("yes"); else Console.WriteLine("no");"#,
        ["yes"]
    };

    expr_operator_bitwise_or_on_flags => {
        r#"struct Bits { public int V; public static Bits operator |(Bits a, Bits b) => new Bits { V = a.V | b.V }; }
Console.WriteLine((new Bits { V = 1 } | new Bits { V = 2 }).V);"#,
        ["3"]
    };

    expr_indexer_class_get_only_int_key => {
        r#"class Bag { int[] data = { 10, 20, 30 }; public int this[int i] => data[i]; }
Console.WriteLine(new Bag()[1]);"#,
        ["20"]
    };

    expr_indexer_class_get_set_int_key => {
        r#"class Buffer { int[] data = new int[3]; public int this[int i] { get => data[i]; set => data[i] = value; } }
var b = new Buffer(); b[2] = 99; Console.WriteLine(b[2]);"#,
        ["99"]
    };

    expr_indexer_class_string_key_get_set => {
        r#"class Map { System.Collections.Generic.Dictionary<string, int> d = new(); public int this[string k] { get => d[k]; set => d[k] = value; } }
var m = new Map(); m["count"] = 7; Console.WriteLine(m["count"]);"#,
        ["7"]
    };

    expr_indexer_struct_get_only => {
        r#"struct Row { int[] cells = { 1, 2, 3 }; public int this[int c] => cells[c]; }
Console.WriteLine(new Row()[0]);"#,
        ["1"]
    };

    expr_indexer_struct_get_set => {
        r#"struct PairStore { int a, b; public int this[int slot] { get => slot == 0 ? a : b; set { if (slot == 0) a = value; else b = value; } } }
var p = new PairStore(); p[0] = 3; p[1] = 9; Console.WriteLine(p[0]); Console.WriteLine(p[1]);"#,
        ["3", "9"]
    };

    expr_indexer_class_computed_from_state => {
        r#"class Scale { int factor = 2; public int this[int input] => input * factor; }
Console.WriteLine(new Scale()[5]);"#,
        ["10"]
    };

    expr_indexer_class_two_int_params => {
        r#"class Grid { int[,] m = { { 1, 2 }, { 3, 4 } }; public int this[int r, int c] => m[r, c]; }
Console.WriteLine(new Grid()[1, 0]);"#,
        ["3"]
    };

    expr_indexer_class_expression_bodied_set_only => {
        r#"class Store { int v; public int Value { get { return v; } set => v = value; } }
var s = new Store(); s.Value = 11; Console.WriteLine(s.Value);"#,
        ["11"]
    };

    expr_indexer_on_nested_static_data_via_instance => {
        r#"class Lookup { int[] table = { 5, 6, 7 }; public int this[int i] => table[i]; }
Console.WriteLine(new Lookup()[2]);"#,
        ["7"]
    };

    expr_constructor_class_expression_body_tuple_assign => {
        r#"class Point { public int X, Y; public Point(int x, int y) => (X, Y) = (x, y); }
var p = new Point(3, 4); Console.WriteLine(p.X); Console.WriteLine(p.Y);"#,
        ["3", "4"]
    };

    expr_constructor_struct_expression_body_single_field => {
        r#"struct Id { public int Value; public Id(int v) => Value = v; }
Console.WriteLine(new Id(42).Value);"#,
        ["42"]
    };

    expr_method_and_property_same_class => {
        r#"class Widget { public int baseVal = 5; public int Base => baseVal; public int Twice() => Base * 2; }
Console.WriteLine(new Widget().Twice());"#,
        ["10"]
    };

    expr_method_property_indexer_combined => {
        r#"class Cache { int[] buf = { 0, 0, 0 }; public int this[int i] { get => buf[i]; set => buf[i] = value; } public int Sum() => buf[0] + buf[1] + buf[2]; }
var c = new Cache(); c[0] = 1; c[1] = 2; c[2] = 3; Console.WriteLine(c.Sum());"#,
        ["6"]
    };

    expr_operator_and_method_on_same_struct => {
        r#"struct Num { public int V; public static Num operator +(Num a, Num b) => new Num { V = a.V + b.V }; public int Double() => V * 2; }
var n = new Num { V = 3 } + new Num { V = 4 }; Console.WriteLine(n.Double());"#,
        ["14"]
    };

    expr_property_decimal_computed => {
        r#"class Price { public decimal unit = 2.5m; public decimal Triple => unit * 3m; }
Console.WriteLine(new Price().Triple);"#,
        ["7.5"]
    };

    expr_method_long_arithmetic => {
        r#"class Wide { public long Add(long a, long b) => a + b; }
Console.WriteLine(new Wide().Add(9000000000L, 1L));"#,
        ["9000000001"]
    };

    expr_indexer_char_key_in_string_map => {
        r#"class CharMap { System.Collections.Generic.Dictionary<char, int> m = new(); public int this[char c] { get => m[c]; set => m[c] = value; } }
var cm = new CharMap(); cm['A'] = 1; Console.WriteLine(cm['A']);"#,
        ["1"]
    };

    expr_method_expression_body_with_conditional => {
        r#"class Sign { public string Label(int n) => n < 0 ? "neg" : n > 0 ? "pos" : "zero"; }
Console.WriteLine(new Sign().Label(-1)); Console.WriteLine(new Sign().Label(0)); Console.WriteLine(new Sign().Label(2));"#,
        ["neg", "zero", "pos"]
    };
}
