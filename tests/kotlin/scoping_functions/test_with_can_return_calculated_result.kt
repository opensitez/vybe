// vybe-test: kotlin/scoping_functions/test_with_can_return_calculated_result
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Range(val start: Int, val end: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val width = with(Range(2, 6)) {
                end - start
            }
            __check((width).toString(), "4")
        }
