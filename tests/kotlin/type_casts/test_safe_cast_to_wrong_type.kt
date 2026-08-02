// vybe-test: kotlin/type_casts/test_safe_cast_to_wrong_type
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 100
            val casted = value as? String
            __check((casted == null).toString(), "true")
        }
