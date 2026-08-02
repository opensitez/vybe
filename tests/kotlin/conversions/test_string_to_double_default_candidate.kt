// vybe-test: kotlin/conversions/test_string_to_double_default_candidate
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun parseOrDefaultDouble(source: String, fallback: Double): Double {
            return try {
                source.toDouble()
            } catch (e: Exception) {
                fallback
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((parseOrDefaultDouble("5.5", 9.5)).toString(), "5.5")
            __check((parseOrDefaultDouble("bad", 9.5)).toString(), "9.5")
        }
