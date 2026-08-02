// vybe-test: kotlin/sealed_types/test_sealed_leafs_can_hold_state_in_constructor
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Packet {
            class Text(val text: String) : Packet()
            class Number(val value: Int) : Packet()
        }

        fun render(packet: Packet): String {
            return when (packet) {
                is Packet.Text -> packet.text
                is Packet.Number -> "n=" + packet.value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(Packet.Text("z"))).toString(), "z")
            __check((render(Packet.Number(6))).toString(), "n=6")
        }
