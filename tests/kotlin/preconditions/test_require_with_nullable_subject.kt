// vybe-test: kotlin/preconditions/test_require_with_nullable_subject
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            val out = try {
                requireNotNull(value)
                "ok"
            } catch (e: IllegalArgumentException) {
                "none"
            }
            __check((out).toString(), "none")
        }
