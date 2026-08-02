// vybe-test: kotlin/builtins/test_floor_ceil_rounding_roundtrip
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3.4
            __check((floor(value) + ceil(value)).toString(), "7")
        }
