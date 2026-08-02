// vybe-test: kotlin/block_expressions/test_for_block_iterator
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun main() { val nums = intArrayOf(1,2,3)
var sum = 0
for (x in run { nums }) { sum += x }
println(sum) }

