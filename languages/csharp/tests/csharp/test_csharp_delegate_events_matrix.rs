use super::helpers::run_csharp;

#[test]
fn delegate_multicast_composition_changes_invocation_sequence() {
    let cases = [
        (3, 2, 3, 11),
        (0, 5, -4, 3),
        (-2, -3, 7, 0),
        (10, 10, 10, 40),
    ];

    for (start, plus_delta, mul_delta, expected_after_both) in cases {
        let expected_after_add_only = start + plus_delta;
        let src = format!(
            r#"
System.Func<int, int> plus = x => x + {plus_delta};
System.Func<int, int> mul = x => x * {mul_delta};
System.Func<int, int> pipeline = plus;
pipeline += mul;
int both = pipeline({start});
pipeline -= mul;
int single = pipeline({start});
Console.WriteLine(both);
Console.WriteLine(single);
"#
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_after_both.to_string(),
                expected_after_add_only.to_string()
            ]
        );
    }
}

#[test]
fn predicate_delegate_applies_to_each_item_matrix() {
    let cases: Vec<(Vec<i32>, i32)> = vec![
        (vec![1, 2, 3, 4], 2),
        (vec![0, -1, 2], 2),
        (vec![], 0),
        (vec![5, 7, 9], 0),
    ];

    for (values, expected) in cases {
        let values_src = values
            .iter()
            .map(i32::to_string)
            .collect::<Vec<_>>()
            .join(", ");
        let src = format!(
            r#"
System.Predicate<int> isPositive = value => value > 0;
int[] values = new int[] {{ {values_src} }};
int positives = 0;
foreach (var value in values) {{
    if (isPositive(value)) positives++;
}}
Console.WriteLine(positives);
"#
        );
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}

#[test]
fn action_handlers_can_be_added_removed_from_events() {
    let src = r#"
class Notifier {
    public event System.Action<int>? Raised;
    public void Raise(int value) => Raised?.Invoke(value);
}

int total = 0;
var notifier = new Notifier();
System.Action<int> add = value => total += value;
System.Action<int> sub = value => total -= value;
notifier.Raised += add;
notifier.Raised += sub;
notifier.Raise(5);
notifier.Raised += add;
notifier.Raise(3);
notifier.Raised -= sub;
notifier.Raise(2);
Console.WriteLine(total);
"#;
assert_eq!(run_csharp(src), vec!["7".to_string()]);
}

#[test]
fn delegate_as_interface_projection_matrix() {
    let pairs = [(1, 2), (3, 0), (-2, 4), (10, 5)];

    for (left, right) in pairs {
        let expected = if right == 0 { left * 10 } else { left / right };
        let src = format!(
            r#"
System.Func<int, int, int> safeDivide = (a, b) => b == 0 ? a * 10 : a / b;
Console.WriteLine(safeDivide({left}, {right}));
"#
        );
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}
