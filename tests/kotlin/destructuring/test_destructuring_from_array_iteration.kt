// vybe-test: kotlin/destructuring/test_destructuring_from_array_iteration
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() { var total = 0
for (point in arrayOf(Pair(2, 3), Pair(5, 6))) { val (x, y) = point
total += x * y }
println(total) }

