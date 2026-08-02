// vybe-test: kotlin/generics/test_generic_function_accepts_nullable_bounded_any
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : Any> ensureNotNull(value: T?): String {
            return value?.toString() ?: "missing"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ensureNotNull(4)).toString(), "4")
            __check((ensureNotNull("z")).toString(), "z")
        }
