// vybe-test: kotlin/annotations/test_annotation_on_top_level_parameter
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED_PARAMETER") fun marker(@Deprecated("x") x: Int): Int { return x + 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((marker(4)).toString(), "6") }
