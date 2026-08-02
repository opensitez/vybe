// vybe-test: kotlin/enums/test_enum_when_expression_returns_defaulted_value
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Level { LOW, MEDIUM, HIGH }

        fun describe(level: Level): String {
            return when (level) {
                Level.LOW -> "low"
                Level.MEDIUM -> "medium"
                Level.HIGH -> "high"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val level = if (2 > 1) Level.HIGH else Level.LOW
            __check((describe(level)).toString(), "high")
            __check((describe(Level.MEDIUM)).toString(), "medium")
        }
