//! Deconstruction of tuples, custom types, and nested structures.
use super::helpers::run_csharp;

#[test]
fn custom_class_with_deconstruct_method_supports_deconstruction() {
    assert_eq!(
        run_csharp(
            r#"class Size {
    public int W, H;
    public void Deconstruct(out int w, out int h) { w=W; h=H; }
}
var (w, h) = new Size{W=3,H=4};
Console.WriteLine(w); Console.WriteLine(h);"#
        ),
        &["3", "4"]
    );
}

#[test]
fn nested_tuple_deconstruction_extracts_inner_elements() {
    assert_eq!(
        run_csharp(
            r#"var ((a,b),(c,d)) = ((1,2),(3,4));
Console.WriteLine(a+b+c+d);"#
        ),
        &["10"]
    );
}

#[test]
fn deconstruction_in_foreach_loop_over_tuple_array() {
    assert_eq!(
        run_csharp(
            r#"var pairs = new[]{(1,"a"),(2,"b"),(3,"c")};
int sum=0;
foreach(var (n, _) in pairs) sum+=n;
Console.WriteLine(sum);"#
        ),
        &["6"]
    );
}

#[test]
fn deconstruction_assignment_to_existing_variables() {
    assert_eq!(
        run_csharp(
            r#"int x=0, y=0;
(x, y) = (5, 10);
Console.WriteLine(x); Console.WriteLine(y);"#
        ),
        &["5", "10"]
    );
}

#[test]
fn record_positional_deconstruct_extracts_all_fields() {
    assert_eq!(
        run_csharp(
            r#"record Point(int X, int Y, int Z);
var p = new Point(1,2,3);
var (x,y,z) = p;
Console.WriteLine(x+y+z);"#
        ),
        &["6"]
    );
}
