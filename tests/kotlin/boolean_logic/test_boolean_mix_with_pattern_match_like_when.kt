// vybe-test: kotlin/boolean_logic/test_boolean_mix_with_pattern_match_like_when
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun main() {
            val values = listOf(true, false, true)
            var trues = 0
            var falses = 0
            for (value in values) {
                when (value) {
                    true -> trues++
                    false -> falses++
                }
            }
            println(trues)
            println(falses)
        }

