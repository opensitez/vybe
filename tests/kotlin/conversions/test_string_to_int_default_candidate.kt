// vybe-test: kotlin/conversions/test_string_to_int_default_candidate
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun parseOrDefault(source: String, fallback: Int): Int {
            return try {
                source.toInt()
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
            __check((parseOrDefault("18", 7)).toString(), "18")
            __check((parseOrDefault("bad", 7)).toString(), "7")
        }
