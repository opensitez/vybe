// vybe-test: kotlin/string_builder_api/test_string_builder_set_length_truncate
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("hello")
            out.setLength(2)
            __check((out.toString()).toString(), "he")
            __check((out.length).toString(), "2")
        }
