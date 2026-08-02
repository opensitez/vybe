// vybe-test: kotlin/data_classes/test_data_class_copy_uses_named_and_positional_semantics_together
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Range(val start: Int, val end: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Range(1, 10)
            val shifted = Range(0, base.end).copy(start = base.start + 1)
            __check((base.toString()).toString(), "Range(start=1, end=10)")
            __check((shifted.toString()).toString(), "Range(start=2, end=10)")
        }
