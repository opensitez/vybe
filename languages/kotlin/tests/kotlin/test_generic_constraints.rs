use crate::helpers::run_prints;

#[test]
fn test_generic_constraints_number_sum_int() {
    let out = run_prints(r#"
        fun <T : Number> add(a: T, b: T): Int = a.toInt() + b.toInt()
        fun main() {
            println(add(1, 2))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_generic_constraints_number_sum_double() {
    let out = run_prints(r#"
        fun <T : Number> sum(a: T, b: T): Int = (a.toDouble() + b.toDouble()).toInt()
        fun main() {
            println(sum(1.2, 3.9))
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_generic_constraints_charsequence_length() {
    let out = run_prints(r#"
        fun <T : CharSequence> len(v: T): Int = v.length
        fun main() {
            println(len("abc"))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_generic_constraints_number_compare() {
    let out = run_prints(r#"
        fun <T : Comparable<T>> top(a: T, b: T): T = if (a > b) a else b
        fun main() {
            println(top(5, 2))
            println(top("k", "a"))
        }
    "#);
    assert_eq!(out, &["5", "k"]);
}

#[test]
fn test_generic_constraints_where_bound_number() {
    let out = run_prints(r#"
        fun <T> score(v: T): Int where T : Number {
            return v.toInt()
        }
        fun main() {
            println(score(7))
            println(score(7.9))
        }
    "#);
    assert_eq!(out, &["7", "7"]);
}

#[test]
fn test_generic_constraints_where_bound_charsequence() {
    let out = run_prints(r#"
        fun <T> first(v: T): Char where T : CharSequence {
            return v.first()
        }
        fun main() {
            println(first("xy"))
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_generic_constraints_dual_bound() {
    let out = run_prints(r#"
        interface Named { val name: String }
        fun <T> label(v: T): String where T : Number, T : Comparable<T> {
            return v.toString()
        }
        fun main() {
            println(label(3))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_generic_constraints_any_to_string() {
    let out = run_prints(r#"
        fun <T> render(v: T): String {
            return v.toString()
        }
        fun main() {
            println(render(true))
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_generic_constraints_invariant_restriction() {
    let out = run_prints(r#"
        fun <T> identity(v: T): T = v
        fun main() {
            println(identity("a"))
            println(identity(12))
        }
    "#);
    assert_eq!(out, &["a", "12"]);
}

#[test]
fn test_generic_constraints_list_of_numbers() {
    let out = run_prints(r#"
        fun <T : Number> total(values: List<T>): Int {
            var n = 0
            for (v in values) n += v.toInt()
            return n
        }
        fun main() {
            println(total(listOf(1, 2, 3)))
            println(total(listOf(1.2, 2.8)))
        }
    "#);
    assert_eq!(out, &["6", "4"]);
}

#[test]
fn test_generic_constraints_list_of_chars() {
    let out = run_prints(r#"
        fun <T : CharSequence> concat(values: List<T>): String = values.joinToString(":")
        fun main() {
            println(concat(listOf("a", "b")))
        }
    "#);
    assert_eq!(out, &["a:b"]);
}

#[test]
fn test_generic_constraints_comparable_max() {
    let out = run_prints(r#"
        fun <T> maxValue(a: T, b: T): T where T : Comparable<T> {
            return if (a >= b) a else b
        }
        fun main() {
            println(maxValue(9, 10))
            println(maxValue("m", "n"))
        }
    "#);
    assert_eq!(out, &["10", "n"]);
}

#[test]
fn test_generic_constraints_charsequence_has_prefix() {
    let out = run_prints(r#"
        fun <T : CharSequence> begins(v: T, prefix: String): Boolean = v.startsWith(prefix)
        fun main() {
            println(begins("hello", "he"))
            println(begins("hello", "x"))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_generic_constraints_number_even() {
    let out = run_prints(r#"
        fun <T : Number> isEven(v: T): Boolean = v.toInt() % 2 == 0
        fun main() {
            println(isEven(4))
            println(isEven(5))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_generic_constraints_number_to_byte() {
    let out = run_prints(r#"
        fun <T : Number> asByte(v: T): Int = v.toByte().toInt()
        fun main() {
            println(asByte(260))
            println(asByte(1.4))
        }
    "#);
    assert_eq!(out, &["4", "1"]);
}

#[test]
fn test_generic_constraints_charsequence_digits_only() {
    let out = run_prints(r#"
        fun <T : CharSequence> isDigits(v: T): Boolean = v.all { it.isDigit() }
        fun main() {
            println(isDigits("1234"))
            println(isDigits("12a4"))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_generic_constraints_mixed_constraints() {
    let out = run_prints(r#"
        fun <T> describe(v: T): String where T : Number, T : Comparable<T> {
            return if (v.toInt() > 4) "big" else "small"
        }
        fun main() {
            println(describe(9))
        }
    "#);
    assert_eq!(out, &["big"]);
}

#[test]
fn test_generic_constraints_numeric_range() {
    let out = run_prints(r#"
        fun <T : Number> between(v: T, min: T, max: T): Boolean {
            val n = v.toInt()
            return n in min.toInt()..max.toInt()
        }
        fun main() {
            println(between(3, 1, 5))
            println(between(9, 1, 5))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_generic_constraints_chain() {
    let out = run_prints(r#"
        fun <T : Number> chain(v: T): String = "n:" + v.toString()
        fun main() {
            println(chain(1))
            println(chain(2.0))
        }
    "#);
    assert_eq!(out, &["n:1", "n:2.0"]);
}

#[test]
fn test_generic_constraints_string_identity() {
    let out = run_prints(r#"
        fun <T> identityString(v: T): String = v.toString()
        fun main() {
            println(identityString(7))
            println(identityString("x"))
        }
    "#);
    assert_eq!(out, &["7", "x"]);
}

#[test]
fn test_generic_constraints_list_join_default() {
    let out = run_prints(r#"
        fun <T> joinOrDash(values: List<T>): String = if (values.isEmpty()) "-" else values.joinToString(",")
        fun main() {
            println(joinOrDash(listOf<Int>()))
            println(joinOrDash(listOf("a", "b")))
        }
    "#);
    assert_eq!(out, &["-", "a,b"]);
}

#[test]
fn test_generic_constraints_pair_compare() {
    let out = run_prints(r#"
        fun <T : Comparable<T>> greater(a: Pair<T, T>): T = if (a.first > a.second) a.first else a.second
        fun main() {
            println(greater(Pair(2, 9)))
            println(greater(Pair("x", "y")))
        }
    "#);
    assert_eq!(out, &["9", "y"]);
}

#[test]
fn test_generic_constraints_restrict_to_nullable_char() {
    let out = run_prints(r#"
        fun <T : CharSequence> safeFirst(v: T?): String = if (v == null || v.isEmpty()) "-" else v[0].toString()
        fun main() {
            println(safeFirst("abc"))
            println(safeFirst(null))
        }
    "#);
    assert_eq!(out, &["a", "-"]);
}

#[test]
fn test_generic_constraints_to_list() {
    let out = run_prints(r#"
        fun <T> toFlat(values: List<List<T>>): List<T> = values.flatten()
        fun main() {
            val out = toFlat(listOf(listOf(1, 2), listOf(3, 4)))
            println(out.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_generic_constraints_minmax_number() {
    let out = run_prints(r#"
        fun <T : Number> minmax(a: T, b: T): String {
            val aInt = a.toInt()
            val bInt = b.toInt()
            return if (aInt <= bInt) "$aInt:$bInt" else "$bInt:$aInt"
        }
        fun main() {
            println(minmax(7, 4))
            println(minmax(2.0, 8.0))
        }
    "#);
    assert_eq!(out, &["4:7", "2:8"]);
}

#[test]
fn test_generic_constraints_compare_length_with_limit() {
    let out = run_prints(r#"
        fun <T : CharSequence> exceeds(v: T, limit: Int): Boolean = v.length > limit
        fun main() {
            println(exceeds("abc", 2))
            println(exceeds("a", 3))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_generic_constraints_sum_int_arrays() {
    let out = run_prints(r#"
        fun <T : Number> sumArray(v: Array<T>): Double {
            var total = 0.0
            for (item in v) total += item.toDouble()
            return total
        }
        fun main() {
            println(sumArray(arrayOf(1, 2, 3)))
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_generic_constraints_reifiable_like_behavior() {
    let out = run_prints(r#"
        fun <T : Number> show(v: T): String = "" + v
        fun main() {
            println(show(1))
            println(show(1.2))
        }
    "#);
    assert_eq!(out, &["1", "1.2"]);
}

#[test]
fn test_generic_constraints_nested_generic_function() {
    let out = run_prints(r#"
        fun <T> outer(v: T): String {
            fun <U : Comparable<U>> inner(a: U, b: U): Boolean = a > b
            return if (inner(v.toString(), "x")) "gt" else "le"
        }
        fun main() {
            println(outer("a"))
            println(outer("z"))
        }
    "#);
    assert_eq!(out, &["le", "gt"]);
}

#[test]
fn test_generic_constraints_number_as_string() {
    let out = run_prints(r#"
        fun <T : Number> asString(v: T): String = v.toString()
        fun main() {
            println(asString(10))
            println(asString(10.8))
        }
    "#);
    assert_eq!(out, &["10", "10.8"]);
}

#[test]
fn test_generic_constraints_charsequence_tail() {
    let out = run_prints(r#"
        fun <T : CharSequence> tail(v: T): String = v.takeLast(1)
        fun main() {
            println(tail("abcd"))
        }
    "#);
    assert_eq!(out, &["d"]);
}

#[test]
fn test_generic_constraints_count_length_if_possible() {
    let out = run_prints(r#"
        fun <T : CharSequence> report(v: T): Int = v.count()
        fun main() {
            println(report("ab"))
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_generic_constraints_pair_of_numbers() {
    let out = run_prints(r#"
        fun <T : Number> totalPair(a: T, b: T): Double = a.toDouble() + b.toDouble()
        fun main() {
            println(totalPair(1, 2))
            println(totalPair(1.5, 2.25))
        }
    "#);
    assert_eq!(out, &["3.0", "3.75"]);
}

#[test]
fn test_generic_constraints_defaulted_callable() {
    let out = run_prints(r#"
        class Box<T>(private val v: T)
        fun <T> valueOrEmpty(v: T?): String = v?.toString() ?: "empty"
        fun main() {
            val a: String? = null
            val b: String? = "x"
            println(valueOrEmpty(a))
            println(valueOrEmpty(b))
        }
    "#);
    assert_eq!(out, &["empty", "x"]);
}
