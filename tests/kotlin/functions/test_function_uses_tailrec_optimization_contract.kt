// vybe-test: kotlin/functions/test_function_uses_tailrec_optimization_contract
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun countdown(start: Int): Int {
            tailrec fun loop(current: Int, acc: Int): Int {
                if (current == 0) return acc
                return loop(current - 1, acc + current)
            }
            return loop(start, 0)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((countdown(4)).toString(), "10")
            __check((countdown(0)).toString(), "0")
        }
