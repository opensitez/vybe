// vybe-test: kotlin/annotations/test_annotation_parsing
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated
        fun oldFunction() {
            __check(("deprecated function executed").toString(), "deprecated function executed")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            oldFunction()
        }
