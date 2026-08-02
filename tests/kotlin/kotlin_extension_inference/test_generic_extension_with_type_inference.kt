// vybe-test: kotlin/kotlin_extension_inference/test_generic_extension_with_type_inference
// origin: languages/kotlin/tests/kotlin/test_kotlin_extension_inference.rs

fun <T : Any> T?.orFallback(default: T): T = this ?: default

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text: String? = null
            val count: Int? = 4
            __check((text.orFallback("x")).toString(), "x")
            __check((count.orFallback(9)).toString(), "4")
        }
