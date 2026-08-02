// vybe-test: kotlin/annotations/test_annotation_array_arguments
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class Labels(vararg val values: String)

        @Labels("a", "b", "c")
        fun report(): String {
            return "counted"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((report()).toString(), "counted")
        }
