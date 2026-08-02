// vybe-test: kotlin/string_builder_api/test_string_builder_append_values
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder()
            out.append("a").append(1).append('b')
            __check((out.toString()).toString(), "a1b")
        }
