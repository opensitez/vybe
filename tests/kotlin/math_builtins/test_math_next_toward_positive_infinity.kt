// vybe-test: kotlin/math_builtins/test_math_next_toward_positive_infinity
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val start = 1.0
            val next = kotlin.math.nextUp(start)
            val moved = next > start
            __check((moved).toString(), "true")
            val down = kotlin.math.nextDown(next)
            __check((down <= next).toString(), "true")
            __check((start == kotlin.math.nextDown(next)).toString(), "true")
        }
