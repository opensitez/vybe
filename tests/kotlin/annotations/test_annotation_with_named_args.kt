// vybe-test: kotlin/annotations/test_annotation_with_named_args
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED_PARAMETER")
        fun log(@Deprecated code: Int): String {
            return code.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((log(7)).toString(), "7")
        }
