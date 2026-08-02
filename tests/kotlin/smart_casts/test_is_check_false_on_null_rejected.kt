// vybe-test: kotlin/smart_casts/test_is_check_false_on_null_rejected
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            __check((value is String).toString(), "false")
        }
