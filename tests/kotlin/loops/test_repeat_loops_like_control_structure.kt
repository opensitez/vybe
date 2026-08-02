// vybe-test: kotlin/loops/test_repeat_loops_like_control_structure
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            repeat(4) {
                total += 1
            }
            __check((total).toString(), "4")
        }
