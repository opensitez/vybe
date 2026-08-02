// vybe-test: kotlin/kotlin_string_line_ops/test_string_sub_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abcdef"
            __check((s.subSequence(1, 4)).toString(), "bcd")
            __check((s.substring(2, 4)).toString(), "cd")
            __check((s.take(2)).toString(), "ab")
            __check((s.drop(2)).toString(), "cdef")
        }
