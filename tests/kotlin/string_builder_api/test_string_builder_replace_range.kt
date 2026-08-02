// vybe-test: kotlin/string_builder_api/test_string_builder_replace_range
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("a-b-c")
            out.replace(1, 2, "B")
            __check((out.toString()).toString(), "aBc")
        }
