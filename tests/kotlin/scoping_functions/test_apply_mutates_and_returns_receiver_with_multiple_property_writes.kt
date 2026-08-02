// vybe-test: kotlin/scoping_functions/test_apply_mutates_and_returns_receiver_with_multiple_property_writes
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Packet {
            var head: String = ""
            var tail: String = ""
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val packet = Packet().apply {
                head = "h"
                tail = "t"
                head += "+"
            }
            __check((packet.head).toString(), "h+")
            __check((packet.tail).toString(), "t")
        }
