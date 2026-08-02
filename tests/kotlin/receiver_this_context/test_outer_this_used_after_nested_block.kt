// vybe-test: kotlin/receiver_this_context/test_outer_this_used_after_nested_block
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Gate {
            val id = "gate"
            inner class Guard {
                fun value(): String {
                    return this@Gate.run { id }
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Gate().Guard().value()).toString(), "gate")
        }
