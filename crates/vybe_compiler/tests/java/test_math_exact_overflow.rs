/// Math exact arithmetic methods detect overflow.
use crate::helpers::run_main;

#[test]
fn math_add_exact_sums_within_range() {
    let out = run_main("System.out.println(Math.addExact(100, 25));");
    assert_eq!(out, vec!["125"]);
}

#[test]
fn math_add_exact_throws_on_integer_overflow() {
    let out = run_main(
        "try { System.out.println(Math.addExact(Integer.MAX_VALUE, 1)); } catch (ArithmeticException e) { System.out.println(\"overflow\"); }",
    );
    assert_eq!(out, vec!["overflow"]);
}

#[test]
fn math_multiply_exact_within_range() {
    let out = run_main("System.out.println(Math.multiplyExact(6, 7));");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn math_multiply_exact_throws_when_product_overflows() {
    let out = run_main(
        "try { System.out.println(Math.multiplyExact(Integer.MAX_VALUE, 2)); } catch (ArithmeticException e) { System.out.println(\"overflow\"); }",
    );
    assert_eq!(out, vec!["overflow"]);
}

#[test]
fn math_subtract_exact_throws_on_underflow() {
    let out = run_main(
        "try { System.out.println(Math.subtractExact(Integer.MIN_VALUE, 1)); } catch (ArithmeticException e) { System.out.println(\"underflow\"); }",
    );
    assert_eq!(out, vec!["underflow"]);
}

#[test]
fn math_negate_exact_throws_on_min_value() {
    let out = run_main(
        "try { System.out.println(Math.negateExact(Integer.MIN_VALUE)); } catch (ArithmeticException e) { System.out.println(\"negate\"); }",
    );
    assert_eq!(out, vec!["negate"]);
}

#[test]
fn math_increment_exact_throws_on_max_value() {
    let out = run_main(
        "try { System.out.println(Math.incrementExact(Integer.MAX_VALUE)); } catch (ArithmeticException e) { System.out.println(\"inc\"); }",
    );
    assert_eq!(out, vec!["inc"]);
}

#[test]
fn math_decrement_exact_throws_on_min_value() {
    let out = run_main(
        "try { System.out.println(Math.decrementExact(Integer.MIN_VALUE)); } catch (ArithmeticException e) { System.out.println(\"dec\"); }",
    );
    assert_eq!(out, vec!["dec"]);
}
