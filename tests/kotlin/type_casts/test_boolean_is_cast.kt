// vybe-test: kotlin/type_casts/test_boolean_is_cast
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val value: Any = true
__check((value is Boolean).toString(), "true") }
