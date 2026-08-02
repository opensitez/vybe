// vybe-test: kotlin/destructuring/test_destructuring_inside_if_with_else
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() { if (true) { val (a, b) = Pair(6, 7)
println(a)
println(b) } else { println("no") } }

