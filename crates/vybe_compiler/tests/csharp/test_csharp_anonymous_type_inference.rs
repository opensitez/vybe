//! Anonymous types infer property names from source expressions.
use super::helpers::run_csharp;

#[test]
fn anonymous_type_infers_property_names_from_simple_identifier_expressions() {
    assert_eq!(
        run_csharp(
            r#"
int width = 4;
string label = "box";
var shape = new { width, label };
Console.WriteLine(shape.width);
Console.WriteLine(shape.label);
"#
        ),
        &["4", "box"]
    );
}
