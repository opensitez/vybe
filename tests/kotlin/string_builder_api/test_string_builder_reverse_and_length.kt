// vybe-test: kotlin/string_builder_api/test_string_builder_reverse_and_length
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("kotlin")
            __check((out.length).toString(), "6")
            out.reverse()
            __check((out.toString()).toString(), "niltok")
            __check((out.length).toString(), "6")
        }
