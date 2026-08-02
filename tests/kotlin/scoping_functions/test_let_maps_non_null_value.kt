// vybe-test: kotlin/scoping_functions/test_let_maps_non_null_value
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5
            val result = value.let { it + 7 }
            __check((result).toString(), "12")
        }
