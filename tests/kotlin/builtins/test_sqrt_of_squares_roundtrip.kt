// vybe-test: kotlin/builtins/test_sqrt_of_squares_roundtrip
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val side = 13.0
            __check((sqrt(side * side)).toString(), "13")
        }
