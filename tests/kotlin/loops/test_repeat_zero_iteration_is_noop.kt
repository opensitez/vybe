// vybe-test: kotlin/loops/test_repeat_zero_iteration_is_noop
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            repeat(0) {
                total += 1
            }
            __check((total).toString(), "0")
        }
