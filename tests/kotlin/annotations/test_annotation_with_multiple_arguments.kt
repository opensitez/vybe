// vybe-test: kotlin/annotations/test_annotation_with_multiple_arguments
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED_VARIABLE", "NAME_SHADOWING")
        fun annotated() {
            __check(("multi").toString(), "multi")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            annotated()
        }
