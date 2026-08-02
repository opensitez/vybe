// vybe-test: kotlin/kotlin_string_line_ops/test_line_sequence_counts
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "x\ny\nz\n"
            val count = value.lineSequence().count()
            val tail = value.lineSequence().last()
            __check((count).toString(), "4")
            __check((tail).toString(), "")
        }
