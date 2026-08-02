// vybe-test: kotlin/string_builder_api/test_string_builder_capacity_growth
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder(4)
            out.append("abcd")
            __check((out.capacity() >= 4).toString(), "true")
            out.append("ef")
            __check((out.capacity() >= 6).toString(), "true")
        }
