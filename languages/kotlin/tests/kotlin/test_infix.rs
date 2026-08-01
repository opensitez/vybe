use crate::helpers::run_prints;

#[test]
fn test_infix_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val p = "key" to "value"
            val (key, value) = p
            println(key)
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["key", "value"]);
}

#[test]
fn test_infix_contains() {
    let out = run_prints(
        r#"
        fun main() {
            if (3 in 1..5) {
                println("inside")
            } else {
                println("outside")
            }
        }
    "#,
    );
    assert_eq!(out, &["inside"]);
}

#[test]
fn test_infix_with_down_to_range() {
    let out = run_prints(
        r#"
        fun main() {
            var sum = 0
            for (i in 5 downTo 1) {
                sum += i
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_custom_infix_function() {
    let out = run_prints(
        r#"
        class Calculator(val base: Int) {
            infix fun plusValue(other: Int): Int {
                return base + other
            }
        }

        fun main() {
            val calc = Calculator(10)
            val res = calc plusValue 5
            println(res)
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_custom_infix_chain() {
    let out = run_prints(
        r#"
        class Counter(val base: Int) {
            infix fun plus(other: Int): Int = base + other
            infix fun minus(other: Int): Int = base - other
        }

        fun main() {
            val c = Counter(10)
            println(c plus 3)
            println(c minus 5)
        }
    "#,
    );
    assert_eq!(out, &["13", "5"]);
}

#[test]
fn test_infix_expression_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val keyValue = "a" to "b" to "c"
            val first = keyValue.first
            val second = keyValue.second
            println(first)
            println(second)
        }
    "#,
    );
    assert_eq!(out, &["a", "b"]);
}

#[test]
fn test_in_operator_on_custom_type() {
    let out = run_prints(
        r#"
        class Bag {
            val values = arrayOf(1, 2, 3)
            fun has(value: Int): Boolean {
                return values[0] == value || values[1] == value || values[2] == value
            }
        }

        fun main() {
            val b = Bag()
            if (2 in 1..4) {
                println("range")
            }
            if (b.has(2)) {
                println("found")
            }
        }
    "#,
    );
    assert_eq!(out, &["range", "found"]);
}

#[test]
fn test_infix_down_to_contains() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            for (n in 4 downTo 1) {
                if (n in 3 downTo 1) {
                    total += n
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_infix_custom_boolean() {
    let out = run_prints(
        r#"
        class Guard {
            infix fun allows(hour: Int): Boolean {
                return hour >= 9 && hour <= 17
            }
        }

        fun main() {
            val shift = Guard()
            println(shift.allows(10))
            println(shift.allows(2))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_to_infix_nested_destructure() {
    let out = run_prints(
        r#"
        fun main() {
            val (left, right) = "left" to "right"
            println(left + "," + right)
        }
    "#,
    );
    assert_eq!(out, &["left,right"]);
}

#[test]
fn test_to_infix_with_numeric_types() {
    let out = run_prints(
        r#"
        fun main() {
            val combo = 2 to 4.5
            val first = combo.first
            val second = combo.second
            println(first)
            println(second)
        }
    "#,
    );
    assert_eq!(out, &["2", "4.5"]);
}

#[test]
fn test_infix_with_step_and_accumulation() {
    let out = run_prints(
        r#"
        fun main() {
            var sum = 0
            for (x in 0..10 step 3) {
                sum += x
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["18"]);
}

#[test]
fn test_infix_with_if_expression() {
    let out = run_prints(
        r#"
        class IntPair(val first: Int, val second: Int) {
            infix fun merge(other: IntPair): Int {
                return (first + second) + (other.first + other.second)
            }
        }

        fun main() {
            val a = IntPair(1, 2)
            val b = IntPair(3, 4)
            println(a merge b)
            println(a.first + a.second)
            println(b.first + b.second)
        }
    "#,
    );
    assert_eq!(out, &["10", "3", "7"]);
}

#[test]
fn test_infix_string_pair_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val pair = "x" to 10
            if (pair.first == "x" && pair.second == 10) {
                println("ok")
            }
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_infix_range_contains() {
    let out = run_prints(
        r#"
fun main() { if (4 in 1..3) { println("yes") } else { println("no") } }
"#,
    );
    assert_eq!(out, &["no"]);
}

#[test]
fn test_infix_not_in_range() {
    let out = run_prints(
        r#"
fun main() { if (5 !in 1..3) { println("no") } else { println("yes") } }
"#,
    );
    assert_eq!(out, &["no"]);
}

#[test]
fn test_infix_down_to_pair() {
    let out = run_prints(
        r#"
fun main() { var sum = 0; for (i in 8 downTo 2) { if (i % 2 == 0) { sum += i } }; println(sum) }
"#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_infix_step_range() {
    let out = run_prints(
        r#"
fun main() { var n = 0; for (v in 0..10 step 4) { n += v }; println(n) }
"#,
    );
    assert_eq!(out, &["18"]);
}

#[test]
fn test_infix_to_construction() {
    let out = run_prints(
        r#"
fun main() { val pair = 10 to 20; println(pair.first); println(pair.second) }
"#,
    );
    assert_eq!(out, &["10", "20"]);
}

#[test]
fn test_infix_to_nested() {
    let out = run_prints(
        r#"
fun main() { val a = (1 to 2); val b = (3 to 4); println(a.first + b.second) }
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_infix_custom_add() {
    let out = run_prints(
        r#"
class Adder(val base: Int) { infix fun plusValue(other: Int): Int = base + other }; fun main() { println(Adder(5) plusValue 4) }
"#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_infix_custom_text() {
    let out = run_prints(
        r#"
class Verb { infix fun shout(other: String): String = other + other }; fun main() { println(Verb() shout "go") }
"#,
    );
    assert_eq!(out, &["gogo"]);
}

#[test]
fn test_infix_contains_on_range() {
    let out = run_prints(
        r#"
fun main() { val r = 1..10; println(4 in r); println(11 in r) }
"#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_infix_negative_range() {
    let out = run_prints(
        r#"
fun main() { val inRange = (-3..3); println(0 in inRange); println(4 in inRange) }
"#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_infix_down_to_then_step() {
    let out = run_prints(
        r#"
fun main() { var total = 0; for (i in 9 downTo 1 step 2) { total += i }; println(total) }
"#,
    );
    assert_eq!(out, &["25"]);
}

#[test]
fn test_infix_chainable_custom() {
    let out = run_prints(
        r#"
class Box(val value: Int) { infix fun plus(other: Box): Int = value + other.value; infix fun minus(other: Box): Int = value - other.value }; fun main() { val a = Box(9); val b = Box(3); println(a plus b); println(a minus b) }
"#,
    );
    assert_eq!(out, &["12", "6"]);
}

#[test]
fn test_infix_with_boolean_result() {
    let out = run_prints(
        r#"
class Window { infix fun contains(value: Int): Boolean { return value % 2 == 0 } }; fun main() { val w = Window(); println(w contains 8); println(w contains 9) }
"#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_infix_in_string_search() {
    let out = run_prints(
        r#"
fun main() { val text = "kotlin"; println("li" in text); println("zz" in text) }
"#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_infix_when_with_in_operator() {
    let out = run_prints(
        r#"
fun score(v: Int): String { return when (v) { in 90..100 -> "A"; in 80..89 -> "B"; else -> "F" } }; fun main() { println(score(95)); println(score(50)) }
"#,
    );
    assert_eq!(out, &["A", "F"]);
}

#[test]
fn test_infix_double_to() {
    let out = run_prints(
        r#"
fun main() { val pair = 2 to 4; println(pair.first + pair.second) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_infix_range_includes_boundaries() {
    let out = run_prints(
        r#"
fun main() { val bounds = 1..5; println(1 in bounds); println(5 in bounds); println(6 in bounds) }
"#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_infix_character_range_membership() {
    let out = run_prints(
        r#"
        fun main() {
            println('b' in 'a'..'c')
            println('z' in 'a'..'c')
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_custom_contains_operator_infix_style() {
    let out = run_prints(
        r#"
        class Window(val min: Int, val max: Int) {
            operator fun contains(value: Int): Boolean = value in min..max
        }

        fun main() {
            val active = Window(1, 5)
            println(3 in active)
            println(7 in active)
            println(8 !in active)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_infix_contains_from_extension() {
    let out = run_prints(
        r#"
        class Window(val values: Set<Int>) {
            operator fun contains(value: Int): Boolean = values.contains(value)
        }

        infix fun Window.has(value: Int): Boolean = value in this

        fun main() {
            val setWindow = Window(setOf(2, 4, 6))
            println(setWindow has 4)
            println(setWindow has 3)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_infix_with_unless_operator_fallback() {
    let out = run_prints(
        r#"
        class Counter(val value: Int) {
            infix fun plus(other: Counter): Int = this.value + other.value
        }

        fun main() {
            val left = Counter(4)
            val right = Counter(6)
            println(left plus right)
            println(right plus left)
        }
    "#,
    );
    assert_eq!(out, &["10", "10"]);
}

#[test]
fn test_infix_multiple_operator_calls_left_to_right() {
    let out = run_prints(
        r#"
        class NumberPair(val value: Int) {
            infix fun add(other: NumberPair): NumberPair = NumberPair(this.value + other.value)
        }

        fun main() {
            val total = NumberPair(1) add NumberPair(2) add NumberPair(3)
            println(total.value)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_infix_to_with_boolean_payload() {
    let out = run_prints(
        r#"
        data class PairBool(val key: String, val enabled: Boolean)

        fun main() {
            val item = "feature" to true
            println(item.first)
            println(item.second)
            val toggled = item.first to (item.second && false)
            println(toggled.second)
        }
    "#,
    );
    assert_eq!(out, &["feature", "true", "false"]);
}

#[test]
fn test_infix_precedence_beats_to_creation_left_side() {
    let out = run_prints(
        r#"
        class Score(val value: Int) {
            infix fun plus(other: Int): Int = value + other
        }

        fun main() {
            val calc = Score(2)
            val total = calc plus 3 * 4
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["14"]);
}

#[test]
fn test_infix_right_operand_is_full_expression() {
    let out = run_prints(
        r#"
        class Score(val value: Int) {
            infix fun plus(other: Int): Int = value + other
        }

        fun main() {
            val calc = Score(2)
            val total = (calc plus 3) * 4
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_infix_contains_respects_short_circuit_in_composite_predicate() {
    let out = run_prints(
        r#"
        class Gate {
            var probes = 0
            operator fun contains(value: Int): Boolean {
                probes += 1
                return value % 2 == 0
            }
        }

        fun main() {
            val gate = Gate()
            val outcome = (1 in gate) || (2 in gate)
            println(outcome)
            println(gate.probes)
        }
    "#,
    );
    assert_eq!(out, &["true", "2"]);
}
