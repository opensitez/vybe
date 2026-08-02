// vybe-test: kotlin/kotlin_string_find_ops/test_string_substring_edges
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abcdef"
            __check((s.substring(1, 3)).toString(), "bc")
            __check((s.substring(2)).toString(), "cdef")
        }
