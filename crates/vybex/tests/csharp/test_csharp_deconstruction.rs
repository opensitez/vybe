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
    tuple_deconstruction_assigns_two_scalars,
    r#"
var (x, y) = (3, 4);
Console.WriteLine(x);
Console.WriteLine(y);
"#,
    ["3", "4"]
);

csharp_case!(
    tuple_deconstruction_swaps_values_via_assignment,
    r#"
int left = 1;
int right = 2;
(left, right) = (right, left);
Console.WriteLine(left);
Console.WriteLine(right);
"#,
    ["2", "1"]
);

csharp_case!(
    tuple_deconstruction_discards_unused_value,
    r#"
var (name, _) = ("Ada", 99);
Console.WriteLine(name);
"#,
    ["Ada"]
);

csharp_case!(
    foreach_deconstruction_reads_tuple_sequence,
    r#"
var pairs = new[] { ("a", 1), ("b", 2) };
foreach (var (letter, number) in pairs) {
    Console.WriteLine(letter + number);
}
"#,
    ["a1", "b2"]
);

csharp_case!(
    deconstruct_method_on_class_assigns_two_values,
    r#"
class Point {
    int x;
    int y;
    public Point(int x, int y) { this.x = x; this.y = y; }
    public void Deconstruct(out int xValue, out int yValue) {
        xValue = x;
        yValue = y;
    }
}
var point = new Point(8, 13);
var (x, y) = point;
Console.WriteLine(x);
Console.WriteLine(y);
"#,
    ["8", "13"]
);

csharp_case!(
    deconstruct_method_returns_three_values,
    r#"
class Color {
    int r;
    int g;
    int b;
    public Color(int r, int g, int b) { this.r = r; this.g = g; this.b = b; }
    public void Deconstruct(out int red, out int green, out int blue) {
        red = r;
        green = g;
        blue = b;
    }
}
var color = new Color(1, 2, 3);
var (red, green, blue) = color;
Console.WriteLine(red + green + blue);
"#,
    ["6"]
);

csharp_case!(
    nested_tuple_deconstruction_reads_inner_values,
    r#"
var ((x, y), label) = ((5, 6), "pt");
Console.WriteLine(label);
Console.WriteLine(x + y);
"#,
    ["pt", "11"]
);

csharp_case!(
    deconstruction_with_existing_variables_reassigns_them,
    r#"
int first = 0;
int second = 0;
(first, second) = (7, 9);
Console.WriteLine(first);
Console.WriteLine(second);
"#,
    ["7", "9"]
);

csharp_case!(
    deconstruction_mixes_string_and_numeric_values,
    r#"
var (name, age) = ("Grace", 42);
Console.WriteLine(name + ":" + age);
"#,
    ["Grace:42"]
);

csharp_case!(
    deconstruction_uses_discards_in_foreach_loop,
    r#"
var items = new[] { ("x", 1), ("y", 2), ("z", 3) };
foreach (var (_, number) in items) {
    Console.WriteLine(number * 10);
}
"#,
    ["10", "20", "30"]
);