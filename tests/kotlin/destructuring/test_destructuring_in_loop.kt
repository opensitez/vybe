// vybe-test: kotlin/destructuring/test_destructuring_in_loop
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() { var sum = 0
for (cell in arrayOf(Pair(1, 2), Pair(4, 5))) { val (x, y) = cell
sum += x + y }
println(sum) }

