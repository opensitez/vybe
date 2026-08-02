// vybe-test: kotlin/throwing_recovery/test_throwing_custom_recovery_chain
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun parseValue(s: String): Int {
            return try {
                s.toInt()
            } catch (e: NumberFormatException) {
                -1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((parseValue("9")).toString(), "9")
            __check((parseValue("bad")).toString(), "-1")
        }
