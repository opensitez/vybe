// vybe-test: kotlin/string_builder_api/test_string_builder_indices_navigation
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("abc")
            __check((out[0]).toString(), "a")
            __check((out[1]).toString(), "b")
            __check((out[2]).toString(), "c")
        }
