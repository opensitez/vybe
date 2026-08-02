// vybe-test: kotlin/conversions/test_string_to_int_failure_path_candidate
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun tryParseOrDefault(value: String): Int {
            return try {
                value.toInt()
            } catch (e: Exception) {
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
            __check((tryParseOrDefault("7")).toString(), "7")
            __check((tryParseOrDefault("bad")).toString(), "-1")
        }
