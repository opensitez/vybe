// vybe-test: kotlin/nullability/test_non_null_assertion_after_check
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val maybe: String? = "ok"
            if (maybe != null) {
                __check((maybe!!).toString(), "ok")
            }
        }
