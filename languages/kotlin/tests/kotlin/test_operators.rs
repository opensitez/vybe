use crate::helpers::run_prints;

#[test]
fn test_custom_plus_minus_operators() {
    let out = run_prints(
        r#"
        class Counter(val value: Int) {
            operator fun plus(other: Counter): Counter = Counter(value + other.value)
            operator fun minus(other: Counter): Counter = Counter(value - other.value)
        }

        fun main() {
            val a = Counter(10)
            val b = Counter(4)
            println((a + b).value)
            println((a - b).value)
        }
    "#,
    );
    assert_eq!(out, &["14", "6"]);
}

#[test]
fn test_custom_unary_operators() {
    let out = run_prints(
        r#"
        class Flag(val value: Int) {
            operator fun unaryMinus(): Flag = Flag(-value)
            operator fun unaryPlus(): Flag = Flag(+value)
        }

        fun main() {
            val value = Flag(8)
            println((-value).value)
            println((+value).value)
        }
    "#,
    );
    assert_eq!(out, &["-8", "8"]);
}

#[test]
fn test_custom_inc_dec_operators() {
    let out = run_prints(
        r#"
        class Counter(var value: Int) {
            operator fun inc(): Counter {
                value += 1
                return this
            }

            operator fun dec(): Counter {
                value -= 1
                return this
            }
        }

        fun main() {
            var counter = Counter(2)
            counter++
            println(counter.value)
            counter--
            println(counter.value)
        }
    "#,
    );
    assert_eq!(out, &["3", "2"]);
}

#[test]
fn test_custom_index_get_set() {
    let out = run_prints(
        r#"
        class Buckets {
            private val data = arrayOf(5, 10, 15)
            operator fun get(index: Int): Int {
                return data[index]
            }
            operator fun set(index: Int, value: Int) {
                data[index] = value
            }
        }

        fun main() {
            val storage = Buckets()
            println(storage[0])
            storage[1] = 25
            println(storage[1])
            println(storage[2])
        }
    "#,
    );
    assert_eq!(out, &["5", "25", "15"]);
}

#[test]
fn test_comparison_operator_custom_type() {
    let out = run_prints(
        r#"
        class Version(val major: Int, val minor: Int) {
            operator fun compareTo(other: Version): Int {
                if (major != other.major) {
                    return major - other.major
                }
                return minor - other.minor
            }
        }

        fun main() {
            val a = Version(1, 4)
            val b = Version(2, 0)
            val c = Version(1, 2)
            println(a < b)
            println(a > c)
            println(a == c)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_range_and_contains_for_custom_window() {
    let out = run_prints(
        r#"
        class Window(val low: Int, val high: Int) {
            operator fun contains(value: Int): Boolean {
                return value >= low && value <= high
            }

            operator fun rangeTo(other: Int): IntRange {
                return low..other
            }
        }

        fun main() {
            val window = Window(1, 4)
            println(2 in window)
            println(6 in window)
            val span = window..5
            var total = 0
            for (value in span) {
                total += value
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "15"]);
}

#[test]
fn test_invoke_operator_call_style() {
    let out = run_prints(
        r#"
        class Transformer {
            operator fun invoke(value: Int): Int {
                return value * value
            }
        }

        fun main() {
            val transform = Transformer()
            println(transform(4))
        }
    "#,
    );
    assert_eq!(out, &["16"]);
}

#[test]
fn test_operator_with_generic_type() {
    let out = run_prints(
        r#"
        class Box<T>(private val value: T) {
            operator fun plus(other: Box<T>): String {
                return this.value.toString() + other.value.toString()
            }
        }

        fun main() {
            println(Box("a") + Box("b"))
            println(Box(1) + Box(2))
        }
    "#,
    );
    assert_eq!(out, &["ab", "12"]);
}

#[test]
fn test_assignable_operator_overloads() {
    let out = run_prints(
        r#"
        class Counter(var value: Int) {
            operator fun plusAssign(other: Counter) {
                value += other.value
            }

            operator fun minusAssign(other: Counter) {
                value -= other.value
            }
        }

        fun main() {
            val acc = Counter(10)
            acc += Counter(3)
            println(acc.value)
            acc -= Counter(1)
            println(acc.value)
        }
    "#,
    );
    assert_eq!(out, &["13", "12"]);
}

#[test]
fn test_numeric_operator_precedence() {
    let out = run_prints(
        r#"
        fun main() {
            println(1 + 2 * 3)
            println((1 + 2) * 3)
            println(10 - 3 * 2)
            println(10 / 2 + 1)
        }
    "#,
    );
    assert_eq!(out, &["7", "9", "4", "6"]);
}

#[test]
fn test_division_and_modulo_with_sign() {
    let out = run_prints(
        r#"
        fun main() {
            println(10 / 3)
            println(10 / 3.0)
            println(10 % 3)
            println(-10 % 3)
        }
    "#,
    );
    assert_eq!(out, &["3", "3.3333333333333335", "1", "-1"]);
}

#[test]
fn test_pre_increment_and_post_increment_difference() {
    let out = run_prints(
        r#"
        fun main() {
            var a = 5
            println(++a)
            println(a)
            println(a++)
            println(a)
        }
    "#,
    );
    assert_eq!(out, &["6", "6", "6", "7"]);
}

#[test]
fn test_compound_assignments_sequence() {
    let out = run_prints(
        r#"
        fun main() {
            var value = 2
            value += 3
            value *= 2
            value -= 1
            value /= 2
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_comparison_chain_is_shortcut() {
    let out = run_prints(
        r#"
        fun main() {
            println(2 + 3 > 3 * 1 && 4 <= 4)
            println(2 + 3 > 3 * 2 || 4 < 1)
            println(5 > 4 && 4 > 3)
            println(5 > 4 && 4 > 3 && 2 > 1)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_bitwise_and_or_xor_and_shift() {
    let out = run_prints(
        r#"
        fun main() {
            println(1 shl 4)
            println(16 shr 2)
            println(5 and 3)
            println(5 or 2)
            println(5 xor 2)
            println(5.inv())
        }
    "#,
    );
    assert_eq!(out, &["16", "4", "1", "7", "7", "-6"]);
}

#[test]
fn test_boolean_unary_ops() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 3
            println(-value)
            println(+value)
            println(!true)
            println(!false)
        }
    "#,
    );
    assert_eq!(out, &["-3", "3", "false", "true"]);
}

#[test]
fn test_short_circuit_avoids_side_effect() {
    let out = run_prints(
        r#"
        var steps = 0

        fun maybeHappens(flag: Boolean): Boolean {
            steps += 1
            return flag
        }

        fun main() {
            println(false && maybeHappens(true))
            println(steps)
            println(true || maybeHappens(false))
            println(steps)
        }
    "#,
    );
    assert_eq!(out, &["false", "0", "true", "0"]);
}

#[test]
fn test_elvis_operator_with_defaults() {
    let out = run_prints(
        r#"
        fun main() {
            val name: String? = null
            val provided: String? = "value"
            println(name ?: "fallback")
            println(provided ?: "fallback")
        }
    "#,
    );
    assert_eq!(out, &["fallback", "value"]);
}

#[test]
fn test_range_contains_and_excludes_end() {
    let out = run_prints(
        r#"
        fun main() {
            val values = 1..5
            println(3 in values)
            println(6 in values)
            println(5 in 1 until 5)
            println(5 !in 1 until 5)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "false", "true"]);
}

#[test]
fn test_down_to_and_step_loop_operator() {
    let out = run_prints(
        r#"
        fun main() {
            var reversed = ""
            var sum = 0
            for (value in 7 downTo 3 step 2) {
                reversed += value.toString()
                sum += value
            }
            println(reversed)
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["753", "15"]);
}

#[test]
fn test_until_range_is_exclusive_end() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            for (value in 1 until 4) {
                total += value
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_string_plus_with_non_string_left() {
    let out = run_prints(
        r#"
        fun main() {
            println(2 + 3 + " apples")
            println("apples " + 2 + 3)
            println("calc " + (2 + 3))
        }
    "#,
    );
    assert_eq!(out, &["5 apples", "apples 23", "calc 5"]);
}

#[test]
fn test_equality_operators_for_data_and_reference() {
    let out = run_prints(
        // `data class`, as the name says: a PLAIN class keeps identity
        // `equals`, so real Kotlin prints false/true/false for this body —
        // only the data modifier makes `a == b` structural (true/true/false).
        r#"
        data class Cell(val value: Int)

        fun main() {
            val a = Cell(1)
            val b = Cell(1)
            val c = a
            println(a == b)
            println(a === c)
            println(a === b)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_mixed_numeric_operand_types() {
    let out = run_prints(
        r#"
        fun main() {
            println(1 + 2.0)
            println(5.5 - 2)
            println(8 / 4.0)
            println(8L / 3L)
        }
    "#,
    );
    // Real Kotlin agrees: `1 + 2.0` and `8 / 4.0` are Double, and Double
    // prints with its decimal point — "3.0" and "2.0". Only the Long/Long
    // division stays integral.
    assert_eq!(out, &["3.0", "3.5", "2.0", "2"]);
}

#[test]
fn test_infered_range_for_loop_upper_bound_expression() {
    let out = run_prints(
        r#"
        fun main() {
            var result = 0
            val multiplier = 1
            for (value in 1..(2 + multiplier)) {
                result += value
            }
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_indexed_get_set_operator_with_custom_index_type() {
    let out = run_prints(
        r#"
        class Slots {
            private val values = arrayOf(9, 8, 7)
            operator fun get(index: Int): Int = values[index]
            operator fun set(index: Int, value: Int) {
                values[index] = value
            }
        }

        fun main() {
            val box = Slots()
            println(box[0] + box[2])
            box[1] = 4
            println(box[1])
        }
    "#,
    );
    assert_eq!(out, &["16", "4"]);
}

#[test]
fn test_in_operator_on_empty_range_and_iterable() {
    let out = run_prints(
        r#"
        fun main() {
            val empty = 1..0
            println(empty.isEmpty())
            println(1 in empty)
            val present = 5 in 1..10
            val absent = 11 in 1..10
            println(present)
            println(absent)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "false"]);
}

#[test]
fn test_arithmetic_left_associativity() {
    let out = run_prints(
        r#"
        fun main() {
            println(10 - 5 - 2)
            println(10 - (5 - 2))
            println(2 + 3 - 1 + 4)
        }
    "#,
    );
    assert_eq!(out, &["3", "7", "8"]);
}

#[test]
fn test_is_and_not_is_checks() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any? = "kotlin"
            println(value is String)
            println(value is Int)
            println(value !is Int)
            val nullValue: Any? = null
            println(nullValue is String)
            println(nullValue !is String)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "false", "true"]);
}

#[test]
fn test_char_range_membership() {
    let out = run_prints(
        r#"
        fun main() {
            val vowels = 'a'..'f'
            println('c' in vowels)
            println('z' in vowels)
            println('a' !in vowels)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "false"]);
}

#[test]
fn test_safe_call_operator_on_nullable_reference() {
    let out = run_prints(
        r#"
        class Holder(val label: String)

        fun main() {
            val absent: Holder? = null
            val present: Holder? = Holder("value")
            println(absent?.label)
            println(present?.label)
        }
    "#,
    );
    assert_eq!(out, &["null", "value"]);
}

#[test]
fn test_elvis_operator_skips_rhs_when_present() {
    let out = run_prints(
        r#"
        var evals = 0

        fun fallback(): Int {
            evals += 1
            return 99
        }

        fun coalesce(value: Int?): Int {
            return value ?: fallback()
        }

        fun main() {
            println(coalesce(12))
            println(evals)
            println(coalesce(null))
            println(evals)
        }
    "#,
    );
    assert_eq!(out, &["12", "0", "99", "1"]);
}

#[test]
fn test_nullable_cast_operator_as_question_mark() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = "kotlin"
            val first: String? = value as? String
            val second: Int? = value as? Int
            val third: Any? = null
            val fourth: String? = third as? String
            println(first)
            println(second)
            println(fourth)
        }
    "#,
    );
    assert_eq!(out, &["kotlin", "null", "null"]);
}

#[test]
fn test_safe_call_contains_with_nullable_progression() {
    let out = run_prints(
        r#"
        fun main() {
            val maybeRange: IntRange? = null
            println((maybeRange?.contains(3)) ?: false)
            val explicit = 1..4
            println(explicit?.contains(3))
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_unsigned_right_shift_operator() {
    let out = run_prints(
        r#"
        fun main() {
            println(-1 ushr 1)
            println(-32 ushr 3)
            println(16 ushr 1)
        }
    "#,
    );
    // Real Kotlin agrees: -32 as u32 is 0xFFFFFFE0; ushr 3 gives 0x1FFFFFFC
    // = 536870908 (536870911 = 0x1FFFFFFF is `-1 ushr 3`, a different input).
    assert_eq!(out, &["2147483647", "536870908", "8"]);
}

#[test]
fn test_custom_contains_with_side_effect() {
    let out = run_prints(
        r#"
        class Gate {
            var probes = 0

            operator fun contains(value: Int): Boolean {
                probes += 1
                return value in 10..20
            }
        }

        fun main() {
            val gate = Gate()
            println(12 in gate)
            println(2 in gate)
            println(gate.probes)
            println((5 in gate) || (12 in gate))
            println(gate.probes)
        }
    "#,
    );
    // Real Kotlin agrees: `(5 in gate)` is false, so `||` MUST evaluate the
    // right side — both probes run and the counter lands on 4, not 3.
    assert_eq!(out, &["true", "false", "2", "true", "4"]);
}

#[test]
fn test_vector_like_operator_overload() {
    let out = run_prints(
        r#"
        class Vector(val x: Int, val y: Int) {
            operator fun times(scale: Int): Vector = Vector(x * scale, y * scale)
            operator fun plus(other: Vector): Vector = Vector(x + other.x, y + other.y)
        }

        fun main() {
            val a = Vector(2, 3)
            val b = a * 4
            val c = b + Vector(1, 1)
            println(b.x)
            println(c.y)
        }
    "#,
    );
    assert_eq!(out, &["8", "13"]);
}

#[test]
fn test_when_is_with_casting_semantics() {
    let out = run_prints(
        r#"
        fun describe(value: Any?): String {
            return when (value) {
                is String -> "str:" + value.length
                is Int -> "int"
                null -> "nil"
                else -> "other"
            }
        }

        fun main() {
            println(describe("kotlin"))
            println(describe(7))
            println(describe(3.14))
            println(describe(null))
        }
    "#,
    );
    assert_eq!(out, &["str:6", "int", "other", "nil"]);
}

#[test]
fn test_not_null_assertion_throws_on_null_reference() {
    let out = run_prints(
        r#"
        fun main() {
            val missing: String? = null
            try {
                println(missing!!)
            } catch (e: NullPointerException) {
                println("null")
            }
        }
    "#,
    );
    assert_eq!(out, &["null"]);
}

#[test]
fn test_logical_operators_preserve_left_to_right_evaluation_for_side_effects() {
    let out = run_prints(
        r#"
        var trace = ""
        fun hit(label: String, value: Boolean): Boolean {
            trace += label
            return value
        }

        fun main() {
            println(hit("a", false) && hit("b", true))
            println(trace)
            trace = ""
            println(hit("c", true) || hit("d", false))
            println(trace)
        }
    "#,
    );
    assert_eq!(out, &["false", "a", "true", "c"]);
}

#[test]
fn test_nested_elvis_chain_and_rhs_evaluation_boundary() {
    let out = run_prints(
        r#"
        var fallbackCalls = 0

        fun fallback(value: String?): String {
            fallbackCalls += 1
            return value ?: "default"
        }

        fun main() {
            val first: String? = null
            val second: String? = null
            val third: String? = "value"
            val present: String? = "keep"
            println(first ?: second ?: fallback(third))
            println(fallbackCalls)
            fallbackCalls = 0
            println(present ?: fallback(present))
            println(fallbackCalls)
        }
    "#,
    );
    assert_eq!(out, &["value", "1", "keep", "0"]);
}

#[test]
fn test_integer_division_by_zero_throws() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println(10 / 0)
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#,
    );
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_floating_division_by_zero_is_infinite() {
    let out = run_prints(
        r#"
        fun main() {
            println(10.0 / 0.0)
            println(-10.0 / 0.0)
            println((0.0 / 0.0).isNaN())
        }
    "#,
    );
    assert_eq!(out, &["Infinity", "-Infinity", "true"]);
}

#[test]
fn test_elvis_with_throw_right_side_only_when_null() {
    let out = run_prints(
        r#"
        fun fail(reason: String): Nothing {
            throw Exception(reason)
        }

        fun main() {
            val value: String? = "ok"
            println(value ?: fail("oops"))
            val missing: String? = null
            try {
                println(missing ?: fail("missing"))
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#,
    );
    assert_eq!(out, &["ok", "caught"]);
}

#[test]
fn test_operator_precedence_with_comparison_and_range() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 4
            println(value + 2 * 3 > 12)
            println((value + 2) * 3 > 12)
            println(value in 1..(2 + 1) * 2)
            println(value in 1..2 + 1 * 2)
        }
    "#,
    );
    // Real Kotlin agrees: additive binds TIGHTER than `..` (rangeExpression
    // is built from additive operands), so `1..2 + 1 * 2` is `1..4` and
    // `4 in 1..4` is true.
    assert_eq!(out, &["false", "true", "true", "true"]);
}

#[test]
fn test_null_coalescing_keeps_original_reference_type() {
    let out = run_prints(
        r#"
        fun main() {
            val source: Any? = "value"
            val text: String = source as? String ?: "fallback"
            println(text)
            val raw: Any? = null
            val again: String = raw as? String ?: "fallback"
            println(again)
        }
    "#,
    );
    assert_eq!(out, &["value", "fallback"]);
}
