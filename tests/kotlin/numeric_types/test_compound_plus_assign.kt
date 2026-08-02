// vybe-test: kotlin/numeric_types/test_compound_plus_assign
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 1
            total += 4
            total += 5
            __check((total).toString(), "10")
            total += -2
            __check((total).toString(), "8")
        }
