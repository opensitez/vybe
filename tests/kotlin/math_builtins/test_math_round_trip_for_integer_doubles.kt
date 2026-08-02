// vybe-test: kotlin/math_builtins/test_math_round_trip_for_integer_doubles
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = floor(5.9)
            val b = ceil(5.1)
            val c = round(5.2)
            __check((a).toString(), "5")
            __check((b).toString(), "6")
            __check((c).toString(), "5")
        }
