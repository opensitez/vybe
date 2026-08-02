// vybe-test: kotlin/annotations/test_annotation_order_is_irrelevant_to_execution
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED")
        @Deprecated("deprecated")
        fun marker(value: Int): Int {
            return value + 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((marker(4)).toString(), "5")
        }
