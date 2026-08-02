// vybe-test: kotlin/block_expressions/test_block_with_early_exit
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun f(): Int { val x = run { return 7 }
println("after")
return x }
fun main() { println(f()) }

