// vybe-test: kotlin/infix/test_infix_down_to_pair
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun main() { var sum = 0
for (i in 8 downTo 2) { if (i % 2 == 0) { sum += i } }
println(sum) }

