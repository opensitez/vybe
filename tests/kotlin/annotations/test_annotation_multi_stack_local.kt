// vybe-test: kotlin/annotations/test_annotation_multi_stack_local
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { @Suppress("UNUSED_PARAMETER") @Deprecated("old") fun local(v: Int): Int { return v + 1 }
__check((local(6)).toString(), "7") }
