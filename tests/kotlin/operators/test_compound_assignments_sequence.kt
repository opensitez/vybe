// vybe-test: kotlin/operators/test_compound_assignments_sequence
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 2
            value += 3
            value *= 2
            value -= 1
            value /= 2
            __check((value).toString(), "4")
        }
