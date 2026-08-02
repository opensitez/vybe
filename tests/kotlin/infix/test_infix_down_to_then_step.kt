// vybe-test: kotlin/infix/test_infix_down_to_then_step
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun main() { var total = 0
for (i in 9 downTo 1 step 2) { total += i }
println(total) }

