//! double.nan, infinities, isNaN/isInfinite/isNegative, division by zero, NaN comparisons.

dart_cases! {
    double_nan_literal_is_nan => {
        r#"void main() {
  print(double.nan.isNaN);
}"#,
        ["true"]
    };

    double_nan_is_not_finite => {
        r#"void main() {
  print(double.nan.isFinite);
}"#,
        ["false"]
    };

    double_infinity_literal => {
        r#"void main() {
  print(double.infinity);
}"#,
        ["Infinity"]
    };

    double_negative_infinity_literal => {
        r#"void main() {
  print(double.negativeInfinity);
}"#,
        ["-Infinity"]
    };

    double_infinity_is_infinite => {
        r#"void main() {
  print(double.infinity.isInfinite);
}"#,
        ["true"]
    };

    double_negative_infinity_is_infinite => {
        r#"void main() {
  print(double.negativeInfinity.isInfinite);
}"#,
        ["true"]
    };

    double_infinity_is_not_nan => {
        r#"void main() {
  print(double.infinity.isNaN);
}"#,
        ["false"]
    };

    double_is_nan_static_on_nan => {
        r#"void main() {
  print(double.isNaN(double.nan));
}"#,
        ["true"]
    };

    double_is_nan_static_on_finite => {
        r#"void main() {
  print(double.isNaN(3.14));
}"#,
        ["false"]
    };

    double_is_infinite_static_on_infinity => {
        r#"void main() {
  print(double.isInfinite(double.infinity));
}"#,
        ["true"]
    };

    double_is_infinite_static_on_finite => {
        r#"void main() {
  print(double.isInfinite(1.0));
}"#,
        ["false"]
    };

    positive_division_by_zero_yields_infinity => {
        r#"void main() {
  print(1.0 / 0.0);
}"#,
        ["Infinity"]
    };

    negative_division_by_zero_yields_negative_infinity => {
        r#"void main() {
  print(-1.0 / 0.0);
}"#,
        ["-Infinity"]
    };

    zero_division_by_zero_yields_nan => {
        r#"void main() {
  print(0.0 / 0.0);
}"#,
        ["NaN"]
    };

    nan_not_equal_to_nan => {
        r#"void main() {
  print(double.nan != double.nan);
}"#,
        ["true"]
    };

    nan_equal_to_nan_is_false => {
        r#"void main() {
  print(double.nan == double.nan);
}"#,
        ["false"]
    };

    nan_less_than_one_is_false => {
        r#"void main() {
  print(double.nan < 1.0);
}"#,
        ["false"]
    };

    nan_greater_than_one_is_false => {
        r#"void main() {
  print(double.nan > 1.0);
}"#,
        ["false"]
    };

    nan_less_than_or_equal_one_is_false => {
        r#"void main() {
  print(double.nan <= 1.0);
}"#,
        ["false"]
    };

    nan_greater_than_or_equal_one_is_false => {
        r#"void main() {
  print(double.nan >= 1.0);
}"#,
        ["false"]
    };

    one_less_than_nan_is_false => {
        r#"void main() {
  print(1.0 < double.nan);
}"#,
        ["false"]
    };

    infinity_greater_than_large_finite => {
        r#"void main() {
  print(double.infinity > 1e308);
}"#,
        ["true"]
    };

    negative_infinity_less_than_large_negative => {
        r#"void main() {
  print(double.negativeInfinity < -1e308);
}"#,
        ["true"]
    };

    infinity_plus_finite_is_infinity => {
        r#"void main() {
  print(double.infinity + 100.0);
}"#,
        ["Infinity"]
    };

    negative_infinity_minus_finite => {
        r#"void main() {
  print(double.negativeInfinity - 50.0);
}"#,
        ["-Infinity"]
    };

    infinity_times_zero_yields_nan => {
        r#"void main() {
  print(double.infinity * 0.0);
}"#,
        ["NaN"]
    };

    infinity_minus_infinity_yields_nan => {
        r#"void main() {
  print(double.infinity - double.infinity);
}"#,
        ["NaN"]
    };

    positive_infinity_is_not_negative => {
        r#"void main() {
  print(double.infinity.isNegative);
}"#,
        ["false"]
    };

    negative_infinity_is_negative => {
        r#"void main() {
  print(double.negativeInfinity.isNegative);
}"#,
        ["true"]
    };

    finite_positive_is_not_negative => {
        r#"void main() {
  print(5.5.isNegative);
}"#,
        ["false"]
    };

    finite_negative_is_negative => {
        r#"void main() {
  print((-2.5).isNegative);
}"#,
        ["true"]
    };

    nan_is_not_negative => {
        r#"void main() {
  print(double.nan.isNegative);
}"#,
        ["false"]
    };

    negative_zero_is_not_negative => {
        r#"void main() {
  print((-0.0).isNegative);
}"#,
        ["false"]
    };

    infinity_equality_with_self => {
        r#"void main() {
  print(double.infinity == double.infinity);
}"#,
        ["true"]
    };

    negative_infinity_equality_with_self => {
        r#"void main() {
  print(double.negativeInfinity == double.negativeInfinity);
}"#,
        ["true"]
    };

    infinity_not_equal_negative_infinity => {
        r#"void main() {
  print(double.infinity != double.negativeInfinity);
}"#,
        ["true"]
    };

    infinity_not_equal_finite => {
        r#"void main() {
  print(double.infinity != 1000.0);
}"#,
        ["true"]
    };

    sqrt_negative_yields_nan => {
        r#"void main() {
  print((-1.0).sqrt());
}"#,
        ["NaN"]
    };

    nan_plus_finite_yields_nan => {
        r#"void main() {
  print(double.nan + 5.0);
}"#,
        ["NaN"]
    };

    nan_times_finite_yields_nan => {
        r#"void main() {
  print(double.nan * 2.0);
}"#,
        ["NaN"]
    };

    infinity_divided_by_infinity_yields_nan => {
        r#"void main() {
  print(double.infinity / double.infinity);
}"#,
        ["NaN"]
    };

    finite_divided_by_infinity_yields_zero => {
        r#"void main() {
  print(42.0 / double.infinity);
}"#,
        ["0.0"]
    };

    finite_divided_by_negative_infinity => {
        r#"void main() {
  print(-42.0 / double.infinity);
}"#,
        ["-0.0"]
    };

    infinity_times_two => {
        r#"void main() {
  print(double.infinity * 2.0);
}"#,
        ["Infinity"]
    };

    negative_infinity_times_two => {
        r#"void main() {
  print(double.negativeInfinity * 2.0);
}"#,
        ["-Infinity"]
    };

    double_max_finite_is_finite => {
        r#"void main() {
  print(double.maxFinite.isFinite);
}"#,
        ["true"]
    };

    double_min_positive_is_finite => {
        r#"void main() {
  print(double.minPositive.isFinite);
}"#,
        ["true"]
    };
}
