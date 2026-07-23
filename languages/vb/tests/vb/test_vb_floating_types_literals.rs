use super::helpers::run_vb;

#[test]
fn float_literal_single() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Single = 1.5\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Single"]
    );
}
#[test]
fn float_literal_single_type_char() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5F\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Single"]
    );
}
#[test]
fn float_literal_single_type_char_legacy() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5!\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Single"]
    );
}

#[test]
fn float_literal_double() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Double = 1.5\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Double"]
    );
}
#[test]
fn float_literal_double_default() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Double"]
    );
}
#[test]
fn float_literal_double_type_char() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5R\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Double"]
    );
}
#[test]
fn float_literal_double_type_char_legacy() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5#\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Double"]
    );
}

#[test]
fn float_literal_decimal() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Decimal = 1.5D\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Decimal"]
    );
}
#[test]
fn float_literal_decimal_type_char() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5D\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Decimal"]
    );
}
#[test]
fn float_literal_decimal_type_char_legacy() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5@\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Decimal"]
    );
}

#[test]
fn float_literal_e_notation_double() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nDim x = 1.5E2\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"),
        vec!["150"]
    );
}
#[test]
fn float_literal_e_notation_double_negative() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nDim x = 1.5E-2\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"),
        vec!["0.015"]
    );
}
#[test]
fn float_literal_e_notation_double_implicit() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 150.0\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Double"]
    );
}
#[test]
fn float_literal_e_notation_single() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1.5E2F\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Single"]
    );
}

#[test]
fn float_nan_double() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Double = Double.NaN\nConsole.WriteLine(Double.IsNaN(x))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_infinity_double() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Double = Double.PositiveInfinity\nConsole.WriteLine(Double.IsInfinity(x))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_negative_infinity_double() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Double = Double.NegativeInfinity\nConsole.WriteLine(x < 0)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_epsilon_double() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Double = Double.Epsilon\nConsole.WriteLine(x > 0)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}

#[test]
fn float_nan_single() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Single = Single.NaN\nConsole.WriteLine(Single.IsNaN(x))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_infinity_single() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Single = Single.PositiveInfinity\nConsole.WriteLine(Single.IsInfinity(x))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}

#[test]
fn float_decimal_precision() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Decimal = 1.0000000000000000000000000001D\nConsole.WriteLine(x > 1D)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_decimal_division() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Decimal = 10D / 3D\nConsole.WriteLine(x > 3.33D)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_double_precision_loss() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Double = 0.1 + 0.2\nConsole.WriteLine(x = 0.3)\nEnd Sub\nEnd Module"
        ),
        vec!["False"]
    );
} // Classic IEEE 754 precision quirk
#[test]
fn float_decimal_precision_exact() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Decimal = 0.1D + 0.2D\nConsole.WriteLine(x = 0.3D)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn float_literal_underscore_separator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1_000.5_5\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["1000.55"]
    );
}
