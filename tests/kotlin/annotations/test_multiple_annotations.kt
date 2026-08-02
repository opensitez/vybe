// vybe-test: kotlin/annotations/test_multiple_annotations
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated
        @Suppress("UNUSED_VARIABLE")
        fun deprecated_function() {
            __check(("deprecated_function").toString(), "deprecated_function")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            deprecated_function()
        }
