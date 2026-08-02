// vybe-test: kotlin/annotations/test_annotation_on_function_overload_target
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED_PARAMETER")
        fun label(a: Int) {
            __check((a).toString(), "9")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            label(9)
        }
