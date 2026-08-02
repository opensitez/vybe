// vybe-test: kotlin/enums/test_enum_long_chain_in_for
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Step { ONE, TWO, THREE, FOUR }
fun main() { var hit = 0
for (s in arrayOf(Step.ONE, Step.TWO, Step.THREE, Step.FOUR)) { hit += 1 }
println(hit) }

