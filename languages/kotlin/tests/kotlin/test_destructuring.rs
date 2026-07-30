use crate::helpers::run_prints;

#[test]
fn test_destructuring_pair_values() {
    let out = run_prints(r#"
        fun makePair(): Pair<Int, String> = Pair(1, "one")

        fun main() {
            val (id, label) = makePair()
            println(id)
            println(label)
        }
    "#);
    assert_eq!(out, &["1", "one"]);
}

#[test]
fn test_destructuring_with_triple() {
    let out = run_prints(r#"
        fun main() {
            val triple = Triple("x", 5, true)
            val (name, count, active) = triple
            println(name)
            println(count)
            println(active)
        }
    "#);
    assert_eq!(out, &["x", "5", "true"]);
}

#[test]
fn test_mutable_destructuring_reassignment() {
    let out = run_prints(r#"
        fun main() {
            var (left, right) = Pair(10, 20)
            left += 1
            val sum = left + right
            println(sum)
        }
    "#);
    assert_eq!(out, &["31"]);
}

#[test]
fn test_destructuring_function_call() {
    let out = run_prints(r#"
        fun splitValue(): Pair<String, Int> = Pair("value", 9)

        fun main() {
            val (text, count) = splitValue()
            println(text + count.toString())
        }
    "#);
    assert_eq!(out, &["value9"]);
}

#[test]
fn test_destructuring_nested_pair() {
    let out = run_prints(r#"
        fun wrap(): Pair<Pair<Int, Int>, Int> = Pair(Pair(3, 4), 5)

        fun main() {
            val (inner, tail) = wrap()
            val (first, second) = inner
            println(first + second + tail)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_destructuring_from_function_call_tuple() {
    let out = run_prints(r#"
        fun coordinates(): Pair<String, Pair<Int, Int>> = Pair("pt", Pair(7, 8))

        fun main() {
            val (name, point) = coordinates()
            val (x, y) = point
            println(name)
            println(x)
            println(y)
        }
    "#);
    assert_eq!(out, &["pt", "7", "8"]);
}

#[test]
fn test_destructuring_with_var_chain() {
    let out = run_prints(r#"
        fun main() {
            var (left, right) = Pair(100, 200)
            left += 10
            right -= 20
            val result = left + right
            println(result)
        }
    "#);
    assert_eq!(out, &["290"]);
}

#[test]
fn test_destructuring_from_local_function_return() {
    let out = run_prints(r#"
        fun coordinates(): Pair<Int, Int> {
            return Pair(9, 7)
        }

        fun main() {
            val (x, y) = coordinates()
            println(x + y)
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_destructuring_with_var_and_reassignment() {
    let out = run_prints(r#"
        fun makePair(): Pair<Int, Int> = Pair(4, 8)

        fun main() {
            var (x, y) = makePair()
            x += 1
            y += 2
            println(x)
            println(y)
        }
    "#);
    assert_eq!(out, &["5", "10"]);
}

#[test]
fn test_destructuring_in_function_arguments() {
    let out = run_prints(r#"
        fun sumPair(pair: Pair<Int, Int>): Int {
            val (left, right) = pair
            return left + right
        }

        fun main() {
            println(sumPair(Pair(10, 15)))
        }
    "#);
    assert_eq!(out, &["25"]);
}

#[test]
fn test_destructuring_triple_and_computation() {
    let out = run_prints(r#"
        fun getTriple(): Triple<Int, Int, Int> = Triple(2, 4, 6)

        fun main() {
            val (a, b, c) = getTriple()
            println(a * b)
            println(c / b)
        }
    "#);
    assert_eq!(out, &["8", "1"]);
}

#[test]
fn test_destructuring_in_for_loop() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (entry in arrayOf(Pair(1, 2), Pair(3, 4), Pair(5, 6))) {
                val (x, y) = entry
                total += x
                total += y
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["21"]);
}

#[test]
fn test_nested_destructuring_on_return() {
    let out = run_prints(r#"
        fun nested(): Pair<Pair<Int, Int>, Int> = Pair(Pair(10, 20), 30)

        fun main() {
            val (pair, tail) = nested()
            val (left, right) = pair
            println(left + right + tail)
        }
    "#);
    assert_eq!(out, &["60"]);
}

#[test]
fn test_destructure_from_array_of_pairs() {
    let out = run_prints(r#"
        fun main() {
            var first = 0
            var second = 0
            for (cell in arrayOf(Pair(3, 4), Pair(1, 2))) {
                val (x, y) = cell
                first += x
                second += y
            }
            println(first)
            println(second)
        }
    "#);
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_destructuring_with_shadowed_vars() {
    let out = run_prints(r#"
        fun main() {
            val pair1 = Pair(1, 2)
            val (a, b) = pair1
            val (c, d) = Pair(a + b, b + 1)
            println(c)
            println(d)
        }
    "#);
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_destructuring_inside_if() {
    let out = run_prints(r#"
        fun main() {
            if (true) {
                val (a, b) = Pair(7, 8)
                println(a)
                println(b)
            } else {
                println("no")
            }
        }
    "#);
    assert_eq!(out, &["7", "8"]);
}

#[test]
fn test_destructuring_with_calculation() {
    let out = run_prints(r#"
        fun main() {
            val (a, b) = Pair(8, 3)
            println(a - b)
            println(a + b)
            println(a * b)
        }
    "#);
    assert_eq!(out, &["5", "11", "24"]);
}

#[test]
fn test_destructuring_pair_simple() {
    let out = run_prints(r#"
fun main() { val pair = Pair(3, 4); val (a, b) = pair; println(a + b) }
"#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_destructuring_triple_simple() {
    let out = run_prints(r#"
fun main() { val t = Triple(1, 2, 3); val (a, b, c) = t; println(a * b * c) }
"#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_destructuring_with_assignment() {
    let out = run_prints(r#"
fun main() { var (left, right) = Pair(12, 3); left = left - 2; println(left * right) }
"#);
    assert_eq!(out, &["30"]);
}

#[test]
fn test_destructuring_function_parameters() {
    let out = run_prints(r#"
fun split(value: Pair<String, Int>): Int { val (label, count) = value; return label.length + count }; fun main() { println(split(Pair("ab", 4)) }
"#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_destructuring_nested_return() {
    let out = run_prints(r#"
fun bundle(): Pair<Pair<Int, Int>, Pair<Int, Int>> = Pair(Pair(1, 2), Pair(3, 4)); fun main() { val (first, second) = bundle(); val (a, b) = first; val (c, d) = second; println(a + b + c + d) }
"#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_destructuring_mutation() {
    let out = run_prints(r#"
var a = 1; var b = 2; fun main() { val pair = Pair(a, b); val (first, second) = pair; println(first + second); }
"#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_destructuring_in_loop() {
    let out = run_prints(r#"
fun main() { var sum = 0; for (cell in arrayOf(Pair(1, 2), Pair(4, 5))) { val (x, y) = cell; sum += x + y }; println(sum) }
"#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_destructuring_inside_if_with_else() {
    let out = run_prints(r#"
fun main() { if (true) { val (a, b) = Pair(6, 7); println(a); println(b) } else { println("no") } }
"#);
    assert_eq!(out, &["6", "7"]);
}

#[test]
fn test_destructuring_triple_derived() {
    let out = run_prints(r#"
fun main() { val (x, y, z) = Triple(4, 5, 6); println(x + y + z) }
"#);
    assert_eq!(out, &["15"]);
}

#[test]
fn test_destructuring_chained_expression() {
    let out = run_prints(r#"
fun makePair(v: Int): Pair<Int, Int> { return Pair(v, v + 1) }; fun main() { val p = makePair(10); val (x, y) = p; println(y - x) }
"#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_destructuring_multiple_calls() {
    let out = run_prints(r#"
fun combine(first: Pair<Int, Int>): Int { val (a, b) = first; return a + b }; fun main() { println(combine(Pair(2, 8))); println(combine(Pair(1, 2))) }
"#);
    assert_eq!(out, &["10", "3"]);
}

#[test]
fn test_destructuring_from_array_iteration() {
    let out = run_prints(r#"
fun main() { var total = 0; for (point in arrayOf(Pair(2, 3), Pair(5, 6))) { val (x, y) = point; total += x * y }; println(total) }
"#);
    assert_eq!(out, &["36"]);
}

#[test]
fn test_destructuring_shadowing_names() {
    let out = run_prints(r#"
fun main() { val (a, b) = Pair(7, 8); val (a2, b2) = Pair(a + 1, b + 1); println(a2); println(b2) }
    "#);
    assert_eq!(out, &["8", "9"]);
}

#[test]
fn test_destructuring_data_class_component_contraction() {
    let out = run_prints(r#"
        data class Point(val x: Int, val y: Int, val z: Int)

        fun origin(): Point = Point(2, 4, 6)

        fun main() {
            val (x, y, z) = origin()
            println(x + y)
            println(z - y)
        }
    "#);
    assert_eq!(out, &["6", "2"]);
}

#[test]
fn test_destructuring_custom_component_functions() {
    let out = run_prints(r#"
        class Holder(private val a: Int, private val b: Int) {
            operator fun component1() = a
            operator fun component2() = b
        }

        fun main() {
            val value = Holder(7, 8)
            val (left, right) = value
            println(left)
            println(right)
            println(left + right)
        }
    "#);
    assert_eq!(out, &["7", "8", "15"]);
}

#[test]
fn test_destructuring_ignores_unused_positions() {
    let out = run_prints(r#"
        fun main() {
            val (first, _, third) = Triple("one", "skip", "three")
            println(first)
            println(third)
        }
    "#);
    assert_eq!(out, &["one", "three"]);
}

#[test]
fn test_destructuring_in_for_each_entry_parameter() {
    let out = run_prints(r#"
        fun main() {
            val pairs = listOf(Pair(1, 2), Pair(3, 4))
            var total = 0
            pairs.forEach { (left, right) ->
                total += left + right
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_destructuring_from_map_entries() {
    let out = run_prints(r#"
        fun main() {
            val inventory = mapOf("apple" to 2, "orange" to 5)
            var labels = ""
            var quantities = 0
            for ((name, count) in inventory) {
                labels += name[0].toString()
                quantities += count
            }
            println(labels)
            println(quantities)
        }
    "#);
    assert_eq!(out, &["ao", "7"]);
}

#[test]
fn test_destructuring_nullable_pair_with_default_fallback() {
    let out = run_prints(r#"
        fun maybe(): Pair<Int, Int>? = if (false) Pair(1, 2) else null

        fun main() {
            val (left, right) = maybe() ?: Pair(0, 0)
            println(left + right)
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_destructuring_iterable_to_variables() {
    let out = run_prints(r#"
        fun main() {
            val source = listOf("a", "b", "c")
            val (first, second, third) = source
            println(first)
            println(second)
            println(third)
        }
    "#);
    assert_eq!(out, &["a", "b", "c"]);
}
