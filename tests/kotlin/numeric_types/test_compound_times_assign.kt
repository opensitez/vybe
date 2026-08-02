// vybe-test: kotlin/numeric_types/test_compound_times_assign
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 2
            total *= 3
            total *= 4
            __check((total).toString(), "24")
            total *= -1
            __check((total).toString(), "-24")
        }
