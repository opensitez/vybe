use crate::helpers::run_prints;

#[test]
fn test_big_integer_add_subtract_multiply() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.math.BigInteger("12345678901234567890")
            val b = java.math.BigInteger("987654321")
            println(a.add(b).toString())
            println(a.subtract(b).toString())
            println(a.multiply(java.math.BigInteger("2")).toString())
        }
    "#,
    );
    assert_eq!(
        out,
        &[
            "12345679888888888891",
            "12345677913580246769",
            "24691357802469135680"
        ]
    );
}

#[test]
fn test_big_integer_pow_mod() {
    let out = run_prints(
        r#"
        fun main() {
            val x = java.math.BigInteger("2")
            val p = x.pow(10)
            println(p.toString())
            val m = p.mod(java.math.BigInteger("1000"))
            println(m.toString())
            println(x.modPow(java.math.BigInteger("5"), java.math.BigInteger("13")).toString())
        }
    "#,
    );
    assert_eq!(out, &["1024", "24", "6"]);
}

#[test]
fn test_big_integer_division_and_remainder() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.math.BigInteger("100")
            val b = java.math.BigInteger("9")
            println(a.divide(b).toString())
            println(a.remainder(b).toString())
            val q = a.divideAndRemainder(b)
            println(q[0].toString())
            println(q[1].toString())
        }
    "#,
    );
    assert_eq!(out, &["11", "1", "11", "1"]);
}

#[test]
fn test_big_decimal_creation_and_arithmetic() {
    let out = run_prints(
        r#"
        import java.math.RoundingMode

        fun main() {
            val a = java.math.BigDecimal("10.5")
            val b = java.math.BigDecimal("4")
            println(a.add(b).toPlainString())
            println(a.subtract(b).toPlainString())
            println(a.multiply(b).toPlainString())
            println(a.divide(b, 2, RoundingMode.HALF_UP).toPlainString())
        }
    "#,
    );
    assert_eq!(out, &["14.5", "6.5", "42.0", "2.62"]);
}

#[test]
fn test_big_decimal_scale_and_precision() {
    let out = run_prints(
        r#"
        import java.math.RoundingMode

        fun main() {
            val value = java.math.BigDecimal("12.3456")
            val reduced = value.setScale(2, RoundingMode.HALF_UP)
            println(reduced.toPlainString())
            println(reduced.scale())
            val up = value.setScale(1, RoundingMode.CEILING)
            println(up.toPlainString())
        }
    "#,
    );
    assert_eq!(out, &["12.35", "2", "12.4"]);
}

#[test]
fn test_big_decimal_compare_and_sign() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.math.BigDecimal("-2")
            val b = java.math.BigDecimal("3")
            println(a.compareTo(b))
            println(a.signum())
            println(b.signum())
            println(java.math.BigDecimal.ZERO.compareTo(java.math.BigDecimal("0.00")))
        }
    "#,
    );
    assert_eq!(out, &["-1", "-1", "1", "0"]);
}

#[test]
fn test_big_decimal_invalid_division() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.math.BigDecimal("10")
            try {
                a.divide(java.math.BigDecimal("0"))
                println("bad")
            } catch (e: ArithmeticException) {
                println(e::class.simpleName)
            }
        }
    "#,
    );
    assert_eq!(out, &["ArithmeticException"]);
}

#[test]
fn test_big_numbers_with_large_exponent_and_abs() {
    let out = run_prints(
        r#"
        fun main() {
            val value = java.math.BigDecimal("-123456")
            println(value.abs().toPlainString())
            val scaled = java.math.BigInteger("2").pow(20)
            println(scaled.toString())
            println(scaled.toString().length)
        }
    "#,
    );
    assert_eq!(out, &["123456", "1048576", "7"]);
}

#[test]
fn test_big_integer_factorial_like_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val one = java.math.BigInteger.ONE
            val two = java.math.BigInteger("2")
            val three = java.math.BigInteger("3")
            val product = one.multiply(two).multiply(three)
            println(product.toString())
            println(product == java.math.BigInteger("6"))
        }
    "#,
    );
    assert_eq!(out, &["6", "true"]);
}

#[test]
fn test_big_decimal_from_long_and_int() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.math.BigDecimal(123L)
            val b = java.math.BigDecimal.valueOf(45L)
            println(a + b)
            println(java.math.BigDecimal("1.5") + java.math.BigDecimal.valueOf(1))
        }
    "#,
    );
    assert_eq!(out, &["168", "2.5"]);
}
