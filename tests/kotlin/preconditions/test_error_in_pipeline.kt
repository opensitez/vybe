// vybe-test: kotlin/preconditions/test_error_in_pipeline
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun parseInt(value: String): Int {
            return value.toIntOrNull() ?: error("invalid")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = try {
                parseInt("x")
            } catch (e: IllegalStateException) {
                -1
            }
            __check((out).toString(), "-1")
        }
