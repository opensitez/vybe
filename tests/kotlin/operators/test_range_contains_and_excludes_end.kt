// vybe-test: kotlin/operators/test_range_contains_and_excludes_end
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = 1..5
            __check((3 in values).toString(), "true")
            __check((6 in values).toString(), "false")
            __check((5 in 1 until 5).toString(), "false")
            __check((5 !in 1 until 5).toString(), "true")
        }
