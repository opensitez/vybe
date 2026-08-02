// vybe-test: kotlin/numeric_types/test_compound_minus_assign
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 20
            total -= 3
            total -= 5
            __check((total).toString(), "12")
            total -= -2
            __check((total).toString(), "14")
        }
