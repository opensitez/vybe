// vybe-test: kotlin/annotations/test_annotation_multi_line_stack
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("legacy")
        @Suppress("UNUSED_VARIABLE")
        fun taggedFunction() {
            __check(("stack").toString(), "stack")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            taggedFunction()
        }
