// vybe-test: kotlin/functions/test_function_local_tailrec_accumulator
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun power(base: Int, exp: Int): Int {
            tailrec fun loop(remaining: Int, acc: Int): Int {
                if (remaining == 0) return acc
                return loop(remaining - 1, acc * base)
            }
            return loop(exp, 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((power(2, 0)).toString(), "1")
            __check((power(3, 3)).toString(), "27")
        }
