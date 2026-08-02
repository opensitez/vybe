// vybe-test: kotlin/math_builtins/test_math_composition_with_min_max
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = max(abs(-12), 8)
            val margin = min(4.7, 9.2)
            __check((score).toString(), "12")
            __check((margin).toString(), "4.7")
            __check((score + margin).toString(), "16.7")
        }
