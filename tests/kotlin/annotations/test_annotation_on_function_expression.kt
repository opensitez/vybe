// vybe-test: kotlin/annotations/test_annotation_on_function_expression
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED") fun action(): String { return "done" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((action()).toString(), "done") }
