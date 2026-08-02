// vybe-test: kotlin/when_expressions/test_when_with_subject_value_block_scope
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun render(level: Int): String {
            return when (level) {
                in 0..9 -> {
                    val label = "low"
                    label + ":" + level
                }
                in 10..19 -> {
                    val offset = level - 10
                    "mid:" + offset
                }
                else -> {
                    val doubled = level * 2
                    "high:" + doubled
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(4)).toString(), "low:4")
            __check((render(13)).toString(), "mid:3")
            __check((render(30)).toString(), "high:60")
        }
