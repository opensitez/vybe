use super::helpers::run_csharp;

fn bool_text(v: bool) -> &'static str {
    if v { "True" } else { "False" }
}

#[test]
fn if_else_chain_matrix() {
    let cases = [(3, 1), (1, 1), (0, 5), (-1, -3), (7, 12), (-4, 10)];

    for (left, right) in cases {
        let src = format!(
            "int left = {left}; int right = {right}; Console.WriteLine(left > right ? 1 : (left == right ? 0 : -1));"
        );
        let expected = if left > right {
            1
        } else if left == right {
            0
        } else {
            -1
        };
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}

#[test]
fn for_loop_even_odd_sum_matrix() {
    let limits = [0, 1, 2, 3, 4, 5, 6, 10];

    for limit in limits {
        let src = format!(
            "int limit = {limit}; int sumEven = 0; int countEven = 0; for (int i = 0; i < limit; i++) {{ if ((i & 1) == 0) {{ sumEven += i; countEven += 1; }} }} Console.WriteLine(sumEven); Console.WriteLine(countEven);"
        );
        let values: Vec<i32> = (0..limit).filter(|value| value % 2 == 0).collect();
        let expected_sum_even: i32 = values.iter().sum();
        let expected_count_even = values.len();
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_sum_even.to_string(),
                expected_count_even.to_string()
            ]
        );
    }
}

#[test]
fn while_loop_halving_and_step_count_matrix() {
    let starts = [1, 2, 3, 4, 5, 8, 15, 32];

    for start in starts {
        let src = format!(
            "int value = {start}; int steps = 0; while (value > 1) {{ value /= 2; steps++; }} Console.WriteLine(steps); Console.WriteLine(value);"
        );
        let mut expected_steps = 0;
        let mut value = start;
        while value > 1 {
            value /= 2;
            expected_steps += 1;
        }
        assert_eq!(
            run_csharp(&src),
            vec![expected_steps.to_string(), value.to_string()]
        );
    }
}

#[test]
fn while_loop_skip_odds_and_sum_matrix() {
    let starts = [0, 1, 2, 3, 4, 5, 7, 10];

    for start in starts {
        let src = format!(
            "int value = {start}; int sum = 0; while (value > 0) {{ if (value % 2 == 0) {{ sum += value; }} value--; }} Console.WriteLine(sum);"
        );
        let expected_sum: i32 = (1..=start).filter(|v| v % 2 == 0).sum();
        assert_eq!(run_csharp(&src), vec![expected_sum.to_string()]);
    }
}

#[test]
fn foreach_matrix_total_and_count() {
    let cases: Vec<Vec<i32>> = vec![
        vec![1, 2, 3],
        vec![5, -5, 7],
        vec![10],
        vec![],
        vec![2, 4, 6, 8],
    ];

    for values in cases {
        let literal = values
            .iter()
            .map(i32::to_string)
            .collect::<Vec<_>>()
            .join(", ");
        let src = format!(
            "var values = new int[] {{ {literal} }}; int sum = 0; int count = 0; foreach (var value in values) {{ sum += value; count++; }} Console.WriteLine(sum); Console.WriteLine(count);"
        );
        let expected_sum: i32 = values.iter().sum();
        assert_eq!(
            run_csharp(&src),
            vec![expected_sum.to_string(), values.len().to_string()]
        );
    }
}

#[test]
fn nested_loop_matrix_dot_product() {
    let sizes = [(1, 1), (2, 3), (3, 4), (5, 0), (6, 2)];

    for (rows, cols) in sizes {
        let src = format!(
            "int rows = {rows}; int cols = {cols}; int sum = 0; int productPairs = 0; for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) {{ sum += r * c; productPairs++; }} Console.WriteLine(sum); Console.WriteLine(productPairs);"
        );
        let expected_sum = (0..rows)
            .flat_map(|r| (0..cols).map(move |c| r * c))
            .sum::<i32>();
        let expected_pairs = rows * cols;
        assert_eq!(
            run_csharp(&src),
            vec![expected_sum.to_string(), expected_pairs.to_string()]
        );
    }
}

#[test]
fn logical_short_circuit_matrix() {
    let cases = [(3, 0), (0, 5), (0, 0), (7, 7), (-1, 2), (-1, 0)];

    for (left, right) in cases {
        let src = format!(
            "int left = {left}; int right = {right}; bool gateAnd = (left != 0) && (right / left > 2); bool gateOr = (left == 0) || (right == 0); Console.WriteLine(gateAnd); Console.WriteLine(gateOr);"
        );
        let expected_gate_and = left != 0 && right / left > 2;
        let expected_gate_or = left == 0 || right == 0;
        assert_eq!(
            run_csharp(&src),
            vec![
                bool_text(expected_gate_and).to_string(),
                bool_text(expected_gate_or).to_string()
            ]
        );
    }
}

#[test]
fn ternary_bucket_matrix() {
    let limits = [0, 1, 2, 3, 4, 5, 6, 10, -1];

    for value in limits {
        let src = format!(
            "int value = {value}; int bucket = value > 5 ? 3 : (value > 0 ? 2 : (value > -1 ? 1 : 0)); Console.WriteLine(bucket);"
        );
        let expected_bucket = if value > 5 {
            3
        } else if value > 0 {
            2
        } else if value > -1 {
            1
        } else {
            0
        };
        assert_eq!(run_csharp(&src), vec![expected_bucket.to_string()]);
    }
}

#[test]
fn pre_post_increment_matrix() {
    let starts = [0, 1, 2, 5, -3, 10];

    for start in starts {
        let src = format!(
            "int value = {start}; int first = value; int second = value++; int third = ++value; Console.WriteLine(first); Console.WriteLine(second); Console.WriteLine(third);"
        );
        let first = start;
        let second = start;
        let third = start + 2;
        assert_eq!(
            run_csharp(&src),
            vec![first.to_string(), second.to_string(), third.to_string()]
        );
    }
}

#[test]
fn modulo_and_division_chain_matrix() {
    let cases = [(10, 3), (12, 5), (-11, 4), (7, 1), (0, 2), (8, 3)];

    for (left, right) in cases {
        let src = format!(
            "int left = {left}; int right = {right}; Console.WriteLine(left / right); Console.WriteLine(left % right);"
        );
        assert_eq!(
            run_csharp(&src),
            vec![(left / right).to_string(), (left % right).to_string(),]
        );
    }
}
