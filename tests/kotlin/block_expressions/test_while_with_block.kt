// vybe-test: kotlin/block_expressions/test_while_with_block
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun main() { var x = 0
while (run { x < 2 }) { x += 1 }
println(x) }

