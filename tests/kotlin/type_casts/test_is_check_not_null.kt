// vybe-test: kotlin/type_casts/test_is_check_not_null
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val value: Any? = "abc"
__check((value is String).toString(), "true") }
