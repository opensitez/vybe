// vybe-test: kotlin/type_casts/test_safe_cast_with_null_source
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Any? = null
            val casted = source as? String
            __check((casted == null).toString(), "true")
        }
