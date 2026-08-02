// vybe-test: kotlin/when_expressions/test_when_subject_evaluates_once_with_side_effects
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

var ticks = 0

        fun next(): Int {
            ticks += 1
            return ticks
        }

        fun classify(): Int {
            return when (next()) {
                1 -> 10
                2 -> 20
                3 -> 30
                else -> 40
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify()).toString(), "10")
            __check((classify()).toString(), "20")
            __check((ticks).toString(), "2")
        }
