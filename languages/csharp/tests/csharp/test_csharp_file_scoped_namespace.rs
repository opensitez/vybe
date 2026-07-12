//! File-scoped namespace declaration syntax (`namespace X;`).

csharp_cases! {
    file_scoped_namespace_class_method_call => {
        r#"namespace Demo;
class Worker { public string Run() => "ok"; }
Console.WriteLine(new Worker().Run());"#,
        ["ok"]
    };

    file_scoped_namespace_struct_field_access => {
        r#"namespace Shapes;
struct Point { public int X; public int Y; }
var p = new Point { X = 2, Y = 3 };
Console.WriteLine(p.X + p.Y);"#,
        ["5"]
    };

    file_scoped_namespace_enum_member => {
        r#"namespace Flags;
enum Mode { Off, On }
Console.WriteLine(Mode.On);"#,
        ["On"]
    };

    file_scoped_namespace_interface_implementation => {
        r#"namespace App;
interface IRun { string Go(); }
class Runner : IRun { public string Go() => "go"; }
IRun r = new Runner();
Console.WriteLine(r.Go());"#,
        ["go"]
    };

    file_scoped_namespace_static_class_method => {
        r#"namespace Tools;
static class MathEx { public static int Double(int n) => n * 2; }
Console.WriteLine(MathEx.Double(6));"#,
        ["12"]
    };

    file_scoped_namespace_nested_class_access => {
        r#"namespace Core;
class Outer { public class Inner { public int Value = 7; } }
Console.WriteLine(new Outer.Inner().Value);"#,
        ["7"]
    };

    file_scoped_namespace_record_positional => {
        r#"namespace Data;
record Pair(int A, int B);
Console.WriteLine(new Pair(1, 2).A);"#,
        ["1"]
    };

    file_scoped_namespace_delegate_invoke => {
        r#"namespace Fn;
delegate int Getter();
Getter g = () => 42;
Console.WriteLine(g());"#,
        ["42"]
    };

    file_scoped_namespace_property_getter => {
        r#"namespace Props;
class Box { public int Size { get; } = 10; }
Console.WriteLine(new Box().Size);"#,
        ["10"]
    };

    file_scoped_namespace_field_mutation => {
        r#"namespace State;
class Counter { public int Count; }
var c = new Counter { Count = 1 };
c.Count = 5;
Console.WriteLine(c.Count);"#,
        ["5"]
    };

    file_scoped_namespace_inheritance => {
        r#"namespace Pets;
class Animal { public string Name = "base"; }
class Dog : Animal { public string Breed = "lab"; }
var d = new Dog();
Console.WriteLine(d.Name); Console.WriteLine(d.Breed);"#,
        ["base", "lab"]
    };

    file_scoped_namespace_generic_class => {
        r#"namespace Gen;
class Box<T> { public T Item; }
var b = new Box<int> { Item = 9 };
Console.WriteLine(b.Item);"#,
        ["9"]
    };

    file_scoped_dotted_namespace_name => {
        r#"namespace Acme.Widgets;
class Widget { public string Name => "w"; }
Console.WriteLine(new Widget().Name);"#,
        ["w"]
    };

    file_scoped_namespace_const_field => {
        r#"namespace Consts;
class Limits { public const int Max = 100; }
Console.WriteLine(Limits.Max);"#,
        ["100"]
    };

    file_scoped_namespace_static_readonly_field => {
        r#"namespace Config;
class App { public static readonly string Env = "prod"; }
Console.WriteLine(App.Env);"#,
        ["prod"]
    };

    file_scoped_namespace_record_struct => {
        r#"namespace Geo;
record struct Point(int X, int Y);
Console.WriteLine(new Point(3, 4).Y);"#,
        ["4"]
    };

    file_scoped_namespace_primary_constructor_class => {
        r#"namespace Svc;
class Service(string name) { public string Name => name; }
Console.WriteLine(new Service("api").Name);"#,
        ["api"]
    };

    file_scoped_namespace_collection_expression => {
        r#"namespace Coll;
int[] data = [1, 2, 3];
Console.WriteLine(data[1]);"#,
        ["2"]
    };

    file_scoped_namespace_local_function => {
        r#"namespace Local;
int Twice(int n) { int Double(int x) => x * 2; return Double(n); }
Console.WriteLine(Twice(5));"#,
        ["10"]
    };

    file_scoped_namespace_switch_expression => {
        r#"namespace Switch;
string Label(int n) => n switch { 1 => "one", _ => "other" };
Console.WriteLine(Label(1));"#,
        ["one"]
    };

    file_scoped_namespace_pattern_is_check => {
        r#"namespace Pat;
object value = "text";
Console.WriteLine(value is string);"#,
        ["True"]
    };

    file_scoped_namespace_sealed_class => {
        r#"namespace Seal;
sealed class Token { public int Id = 1; }
Console.WriteLine(new Token().Id);"#,
        ["1"]
    };

    file_scoped_namespace_abstract_class_with_derived => {
        r#"namespace Abs;
abstract class Base { public abstract int Get(); }
class Impl : Base { public override int Get() => 4; }
Console.WriteLine(new Impl().Get());"#,
        ["4"]
    };

    file_scoped_namespace_typeof_name => {
        r#"namespace Reflect;
class Item { }
Console.WriteLine(typeof(Item).Name);"#,
        ["Item"]
    };

    file_scoped_namespace_nameof_type => {
        r#"namespace Names;
class Widget { }
Console.WriteLine(nameof(Widget));"#,
        ["Widget"]
    };

    file_scoped_namespace_multiple_types_same_file => {
        r#"namespace Duo;
class A { public int Value = 1; }
class B { public int Value = 2; }
Console.WriteLine(new A().Value); Console.WriteLine(new B().Value);"#,
        ["1", "2"]
    };

    file_scoped_namespace_static_property => {
        r#"namespace Static;
class Cache { public static int Size { get; set; } = 5; }
Console.WriteLine(Cache.Size);"#,
        ["5"]
    };

    file_scoped_namespace_method_with_params => {
        r#"namespace Calc;
class Adder { public int Sum(int a, int b) => a + b; }
Console.WriteLine(new Adder().Sum(2, 3));"#,
        ["5"]
    };

    file_scoped_namespace_operator_overload => {
        r#"namespace Ops;
struct V { public int N; public static V operator +(V a, V b) => new V { N = a.N + b.N }; }
var r = new V { N = 2 } + new V { N = 3 };
Console.WriteLine(r.N);"#,
        ["5"]
    };

    file_scoped_namespace_init_property_initializer => {
        r#"namespace Init;
class Config { public int Port { get; init; } = 80; }
Console.WriteLine(new Config { Port = 443 }.Port);"#,
        ["443"]
    };

    file_scoped_namespace_required_property => {
        r#"namespace Req;
class User { public required string Name { get; set; } }
Console.WriteLine(new User { Name = "Ada" }.Name);"#,
        ["Ada"]
    };

    file_scoped_namespace_extension_method_holder => {
        r#"namespace Ext;
static class IntExt { public static int Inc(this int n) => n + 1; }
Console.WriteLine(4.Inc());"#,
        ["5"]
    };

    file_scoped_namespace_lambda_in_method => {
        r#"namespace Lambda;
class Fn { public int Run() { System.Func<int, int> f = x => x + 1; return f(3); } }
Console.WriteLine(new Fn().Run());"#,
        ["4"]
    };

    file_scoped_namespace_try_catch => {
        r#"namespace Err;
string Read() { try { return "ok"; } catch { return "bad"; } }
Console.WriteLine(Read());"#,
        ["ok"]
    };

    file_scoped_namespace_string_interpolation => {
        r#"namespace Text;
class Tag { public string Label(string name) => $"hi {name}"; }
Console.WriteLine(new Tag().Label("Ann"));"#,
        ["hi Ann"]
    };

    file_scoped_namespace_array_length => {
        r#"namespace Arr;
int[] nums = new[] { 1, 2, 3 };
Console.WriteLine(nums.Length);"#,
        ["3"]
    };

    file_scoped_namespace_list_generic => {
        r#"namespace Lists;
var list = new System.Collections.Generic.List<int> { 1, 2 };
Console.WriteLine(list.Count);"#,
        ["2"]
    };

    file_scoped_namespace_deep_namespace_path => {
        r#"namespace A.B.C.D;
class Node { public string Name => "deep"; }
Console.WriteLine(new Node().Name);"#,
        ["deep"]
    };

    file_scoped_namespace_public_class_from_same_file => {
        r#"namespace Pub;
public class Visible { public string Text => "seen"; }
Console.WriteLine(new Visible().Text);"#,
        ["seen"]
    };

    file_scoped_namespace_readonly_struct => {
        r#"namespace Immut;
readonly struct Pair { public readonly int A; public readonly int B; public Pair(int a, int b) { A = a; B = b; } }
Console.WriteLine(new Pair(2, 3).A + new Pair(2, 3).B);"#,
        ["5"]
    };
}
