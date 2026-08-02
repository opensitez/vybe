// vybe-test: kotlin/annotations/test_file_annotation_target_is_parsed
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@file:Suppress("UNUSED_VARIABLE")
        fun ignored() {}

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("ok").toString(), "ok")
        }
