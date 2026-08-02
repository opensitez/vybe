// vybe-test: kotlin/builtins/test_floor_and_ceil_for_integer_input
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((floor(4.0)).toString(), "4")
            __check((ceil(4.0)).toString(), "4")
        }
