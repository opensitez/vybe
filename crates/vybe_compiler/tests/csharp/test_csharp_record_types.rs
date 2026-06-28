//! Record semantics: positional, nominal, equality, ToString, Deconstruct, inheritance.
use super::helpers::run_csharp;

#[test]
fn positional_record_constructor_sets_properties() {
    assert_eq!(
        run_csharp(
            r#"record Point(int X, int Y); var p = new Point(3,4); Console.WriteLine(p.X); Console.WriteLine(p.Y);"#
        ),
        &["3", "4"]
    );
}

#[test]
fn records_generated_equals_compares_all_properties() {
    assert_eq!(
        run_csharp(
            r#"record Point(int X, int Y);
var a = new Point(1,2); var b = new Point(1,2); var c = new Point(1,3);
Console.WriteLine(a == b);
Console.WriteLine(a == c);"#
        ),
        &["True", "False"]
    );
}

#[test]
fn records_generated_tostring_includes_property_names_and_values() {
    assert_eq!(
        run_csharp(
            r#"record Tag(string Name);
Console.WriteLine(new Tag("admin").ToString().Contains("admin"));"#
        ),
        &["True"]
    );
}

#[test]
fn positional_record_supports_deconstruction() {
    assert_eq!(
        run_csharp(
            r#"record Size(int W, int H);
var s = new Size(10,20);
var (w,h) = s;
Console.WriteLine(w); Console.WriteLine(h);"#
        ),
        &["10", "20"]
    );
}

#[test]
fn with_expression_leaves_original_unchanged() {
    assert_eq!(
        run_csharp(
            r#"record Point(int X, int Y);
var p = new Point(1,2);
var q = p with { X=9 };
Console.WriteLine(p.X); Console.WriteLine(q.X);"#
        ),
        &["1", "9"]
    );
}

#[test]
fn nominal_record_with_init_properties() {
    assert_eq!(
        run_csharp(
            r#"record Config { public string Host { get; init; } public int Port { get; init; } }
var c = new Config { Host="localhost", Port=8080 };
Console.WriteLine(c.Host); Console.WriteLine(c.Port);"#
        ),
        &["localhost", "8080"]
    );
}

#[test]
fn record_struct_is_a_value_type() {
    assert_eq!(
        run_csharp(
            r#"record struct Coord(double Lat, double Lon);
var a = new Coord(1.0, 2.0);
var b = a;
Console.WriteLine(a == b);"#
        ),
        &["True"]
    );
}

#[test]
fn record_inheritance_shares_base_properties() {
    assert_eq!(
        run_csharp(
            r#"record Animal(string Name);
record Dog(string Name, string Breed) : Animal(Name);
var d = new Dog("Rex","Lab");
Console.WriteLine(d.Name); Console.WriteLine(d.Breed);"#
        ),
        &["Rex", "Lab"]
    );
}

#[test]
fn record_hash_code_equal_for_equal_instances() {
    assert_eq!(
        run_csharp(
            r#"record Tag(string Name);
var a = new Tag("x"); var b = new Tag("x");
Console.WriteLine(a.GetHashCode() == b.GetHashCode());"#
        ),
        &["True"]
    );
}
