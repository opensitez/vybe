// vybe-test: kotlin/reified_generics/test_reified_nullable_check
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> safeCast(value: Any?): Boolean = value is T

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: String? = null
            val b: String? = "x"
            __check((safeCast<String?>(a)).toString(), "true")
            __check((safeCast<String?>(b)).toString(), "true")
        }
