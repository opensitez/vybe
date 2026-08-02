// vybe-test: kotlin/annotations/test_annotation_with_arguments
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNCHECKED_CAST")
        fun castFunc() {
            __check(("annotated with args").toString(), "annotated with args")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            castFunc()
        }
