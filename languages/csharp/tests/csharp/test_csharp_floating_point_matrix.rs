use super::helpers::run_csharp;

#[test]
fn double_matrix_sign_and_abs() {
    let values: Vec<f64> = vec![0.0, 1.0, -1.0, 12.0, -25.0, 3.5, -4.25];

    for value in values {
        let expected_sign = if value > 0.0 {
            1
        } else if value < 0.0 {
            -1
        } else {
            0
        };
        let expected_abs = value.abs().trunc() as i64;
        let src = format!(
            "double value = {value}; Console.WriteLine(Math.Sign(value)); Console.WriteLine((long)Math.Abs(value));"
        );
        assert_eq!(
            run_csharp(&src),
            vec![expected_sign.to_string(), expected_abs.to_string()]
        );
    }
}

#[test]
fn double_matrix_floor_and_ceiling() {
    let values: Vec<f64> = vec![-3.2, -0.5, 0.2, 1.5, 2.99, 4.0];

    for value in values {
        let src = format!(
            "double value = {value}; Console.WriteLine((long)Math.Floor(value)); Console.WriteLine((long)Math.Ceiling(value));"
        );
        let expected_floor = value.floor() as i64;
        let expected_ceil = value.ceil() as i64;
        assert_eq!(
            run_csharp(&src),
            vec![expected_floor.to_string(), expected_ceil.to_string()]
        );
    }
}

#[test]
fn double_matrix_pow_and_sqrt_roundtrip() {
    let values: [f64; 7] = [0.0, 1.0, 2.0, 3.0, 4.0, 10.0, 12.0];

    for value in values {
        let src = format!(
            "double value = {value}; double squared = Math.Pow(value, 2.0); double sqrt = Math.Sqrt(squared); Console.WriteLine((long)squared); Console.WriteLine((long)sqrt);"
        );
        let expected_sqrt = value * value;
        assert_eq!(
            run_csharp(&src),
            vec![
                (expected_sqrt.trunc() as i64).to_string(),
                (expected_sqrt.sqrt() as i64).to_string()
            ]
        );
    }
}

#[test]
fn double_matrix_min_max_pairs() {
    let pairs: [(f64, f64); 6] = [
        (1.0, 2.0),
        (2.5, -1.0),
        (-3.0, -8.0),
        (0.0, 0.0),
        (100.0, 50.0),
        (-12.0, 4.0),
    ];

    for (left, right) in pairs {
        let src = format!(
            "double left = {left}; double right = {right}; Console.WriteLine(Math.Min(left, right)); Console.WriteLine(Math.Max(left, right));"
        );
        let expected_min = left.min(right);
        let expected_max = left.max(right);
        assert_eq!(
            run_csharp(&src),
            vec![expected_min.to_string(), expected_max.to_string()]
        );
    }
}

#[test]
fn double_matrix_truncate_and_round() {
    let values: [f64; 7] = [-4.9, -4.1, -3.0, 0.0, 2.2, 4.8, 6.0];

    for value in values {
        let src = format!(
            "double value = {value}; Console.WriteLine((long)Math.Truncate(value)); Console.WriteLine((long)Math.Round(value));"
        );
        let expected_trunc = value.trunc() as i64;
        let expected_round = value.round() as i64;
        assert_eq!(
            run_csharp(&src),
            vec![expected_trunc.to_string(), expected_round.to_string()]
        );
    }
}

#[test]
fn double_matrix_division_and_remainder() {
    let cases: [(i64, i64); 6] = [(10, 2), (9, 3), (12, 5), (7, 2), (100, 6), (45, 8)];

    for (left, right) in cases {
        let left_as_double = left as f64;
        let right_as_double = right as f64;
        let src = format!(
            "double left = {left_as_double}; double right = {right_as_double}; Console.WriteLine((long)(left / right)); Console.WriteLine((long)(left % right));"
        );
        let expected_div = left / right;
        let expected_mod = left % right;
        assert_eq!(
            run_csharp(&src),
            vec![expected_div.to_string(), expected_mod.to_string()]
        );
    }
}

#[test]
fn double_matrix_scale_and_offset_matrix() {
    let cases: [(f64, f64, f64); 5] = [
        (-3.0, 2.0, 1.0),
        (4.0, -1.0, 3.0),
        (0.0, 5.0, 10.0),
        (2.5, 2.0, -1.0),
        (-1.5, 3.0, 4.0),
    ];

    for (base, scale, offset) in cases {
        let src = format!(
            "double baseValue = {base}; double scaled = baseValue * {scale} + {offset}; Console.WriteLine((long)scaled);"
        );
        let expected_scaled = (base * scale + offset).trunc() as i64;
        assert_eq!(run_csharp(&src), vec![expected_scaled.to_string()]);
    }
}
