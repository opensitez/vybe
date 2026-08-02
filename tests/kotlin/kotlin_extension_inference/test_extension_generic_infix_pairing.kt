// vybe-test: kotlin/kotlin_extension_inference/test_extension_generic_infix_pairing
// origin: languages/kotlin/tests/kotlin/test_kotlin_extension_inference.rs

fun <T> T.thenValue(v: T): List<T> = listOf(this, v)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1.thenValue(2)).toString(), "[1, 2]")
            __check(("a".thenValue("b")).toString(), "[a, b]")
        }
