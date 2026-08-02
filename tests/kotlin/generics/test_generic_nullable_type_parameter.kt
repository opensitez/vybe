// vybe-test: kotlin/generics/test_generic_nullable_type_parameter
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> unwrapOrNone(value: T?): String {
            return value?.toString() ?: "none"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((unwrapOrNone("ok")).toString(), "ok")
            __check((unwrapOrNone(null as String?)).toString(), "none")
            __check((unwrapOrNone(0)).toString(), "0")
        }
