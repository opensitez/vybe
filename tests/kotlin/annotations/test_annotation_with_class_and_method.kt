// vybe-test: kotlin/annotations/test_annotation_with_class_and_method
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("old") class Marker { @Suppress("UNUSED_PARAMETER") fun tagged(@Deprecated("id") id: Int): Int { return id } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Marker().tagged(11)).toString(), "11") }
