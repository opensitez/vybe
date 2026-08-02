// vybe-test: kotlin/generics/test_generic_reified_like_inference_from_arguments
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> asStringList(value: T, converter: (T) -> String): String {
            return converter(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asStringList(3, { it.toString() })).toString(), "3")
            __check((asStringList(false, { v -> v.toString() })).toString(), "false")
        }
