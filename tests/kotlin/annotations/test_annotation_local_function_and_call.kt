// vybe-test: kotlin/annotations/test_annotation_local_function_and_call
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { @Deprecated("local") fun local() { __check(("local").toString(), "local") }
local() }
