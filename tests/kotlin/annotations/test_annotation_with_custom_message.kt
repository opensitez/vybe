// vybe-test: kotlin/annotations/test_annotation_with_custom_message
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED_PARAMETER") fun tagged(@Suppress("UNUSED") x: Int, @Deprecated("bad") y: Int): Int { return x + y }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((tagged(2, 3)).toString(), "5") }
