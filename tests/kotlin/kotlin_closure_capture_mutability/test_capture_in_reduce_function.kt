// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_in_reduce_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var marker = ""
            val values = listOf("a", "b")
            values.reduce { acc, item ->
                marker = acc + item
                marker
            }
            __check((marker).toString(), "ab")
        }
