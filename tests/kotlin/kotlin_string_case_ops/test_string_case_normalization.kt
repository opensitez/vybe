// vybe-test: kotlin/kotlin_string_case_ops/test_string_case_normalization
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_case_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "HeLLo"
            __check((value.lowercase()).toString(), "hello")
            __check((value.uppercase()).toString(), "HELLO")
        }
