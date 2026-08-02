// vybe-test: kotlin/annotations/test_annotation_on_object_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Sender {
    companion object {
        @Deprecated("legacy")
        @Suppress("UNUSED")
        val token = "go"

        @Suppress("UNUSED")
        fun code(): String = token
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    __check((Sender.code()).toString(), "go")
}
