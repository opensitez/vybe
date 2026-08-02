// vybe-test: kotlin/type_casts/test_nullable_is_nullable_type
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            __check((value is String?).toString(), "true")
            __check((value is String).toString(), "false")
        }
