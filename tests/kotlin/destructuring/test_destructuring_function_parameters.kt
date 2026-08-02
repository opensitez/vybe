// vybe-test: kotlin/destructuring/test_destructuring_function_parameters
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun split(value: Pair<String, Int>): Int { val (label, count) = value
return label.length + count }
fun main() { println(split(Pair("ab", 4)) }

