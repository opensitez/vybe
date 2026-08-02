// vybe-test: kotlin/annotations/test_annotation_on_extension_receiver
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("legacy") fun String.wrap(): String = "[" + this + "]"
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check(("ok".wrap()).toString(), "[ok]") }
