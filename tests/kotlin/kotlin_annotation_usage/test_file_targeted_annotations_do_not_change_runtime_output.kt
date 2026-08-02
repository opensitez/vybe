// vybe-test: kotlin/kotlin_annotation_usage/test_file_targeted_annotations_do_not_change_runtime_output
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.FILE)
        annotation class FileMeta

        @FileMeta
        fun marker() = true

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((marker()).toString(), "true")
        }
