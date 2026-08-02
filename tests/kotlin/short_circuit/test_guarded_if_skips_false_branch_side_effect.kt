// vybe-test: kotlin/short_circuit/test_guarded_if_skips_false_branch_side_effect
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = 0
            val a = true
            if (a && false) {
                log += 1
            } else {
                log += 2
            }
            __check((log).toString(), "2")
        }
