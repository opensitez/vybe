// vybe-test: kotlin/type_casts/test_safe_cast_success
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "hello"
            val casted = value as? String
            __check((casted!!).toString(), "hello")
        }
