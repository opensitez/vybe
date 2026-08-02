// vybe-test: kotlin/infix/test_infix_step_range
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun main() { var n = 0
for (v in 0..10 step 4) { n += v }
println(n) }

