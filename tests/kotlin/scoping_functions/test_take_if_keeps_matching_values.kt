// vybe-test: kotlin/scoping_functions/test_take_if_keeps_matching_values
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((10.takeIf { it > 5 }).toString(), "10")
            __check((10.takeIf { it % 2 == 0 }).toString(), "10")
        }
