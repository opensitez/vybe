// vybe-test: kotlin/preconditions/test_require_not_null_non_null
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = "ok"
            __check((requireNotNull(value)).toString(), "ok")
        }
