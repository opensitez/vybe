// vybe-test: kotlin/sealed_types/test_sealed_branches_can_be_mapped_without_else
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Packet {
            class Text : Packet()
            class Number : Packet()
        }

        fun describe(packet: Packet): String {
            return when (packet) {
                is Packet.Text -> "text"
                is Packet.Number -> "number"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Packet.Text())).toString(), "text")
            __check((describe(Packet.Number())).toString(), "number")
        }
