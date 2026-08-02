// vybe-test: kotlin/numeric_types/test_compound_divide_assign_integer
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 32
            total /= 4
            total /= 2
            __check((total).toString(), "4")
        }
