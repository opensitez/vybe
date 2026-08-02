// vybe-test: kotlin/annotations/test_annotation_extension_receiver
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("old")
        fun String.highlight(): String {
            return "<<" + this + ">>"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("ok".highlight()).toString(), "<<ok>>")
        }
