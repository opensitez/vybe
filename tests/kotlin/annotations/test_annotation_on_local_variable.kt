// vybe-test: kotlin/annotations/test_annotation_on_local_variable
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            @Deprecated("temp")
            val status = "pending"
            __check((status).toString(), "pending")
        }
