// vybe-test: kotlin/control_flow/test_if_expression_skips_false_branch_side_effects
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

var hits = 0

        fun bump(): Int {
            hits += 1
            return 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = if (1 == 1) {
                7
            } else {
                bump()
            }
            __check((value).toString(), "7")
            __check((hits).toString(), "0")
        }
