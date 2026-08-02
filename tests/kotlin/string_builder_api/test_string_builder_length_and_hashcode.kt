// vybe-test: kotlin/string_builder_api/test_string_builder_length_and_hashcode
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("hash")
            __check((out.length).toString(), "4")
            __check((out.toString().hashCode() == out.hashCode()).toString(), "false")
        }
