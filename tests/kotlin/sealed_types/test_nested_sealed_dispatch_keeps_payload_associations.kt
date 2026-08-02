// vybe-test: kotlin/sealed_types/test_nested_sealed_dispatch_keeps_payload_associations
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Packet {
            class Left(val code: Int) : Packet()
            class Right(val label: String) : Packet()
        }

        sealed class Wrapper {
            class Item(val packet: Packet) : Wrapper()
            class Empty : Wrapper()
        }

        fun describe(wrapper: Wrapper): String {
            return when (wrapper) {
                is Wrapper.Empty -> "none"
                is Wrapper.Item -> when (wrapper.packet) {
                    is Packet.Left -> "L" + wrapper.packet.code
                    is Packet.Right -> "R" + wrapper.packet.label
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
            __check((describe(Wrapper.Item(Packet.Left(2)))).toString(), "L2")
            __check((describe(Wrapper.Item(Packet.Right("ok")))).toString(), "Rok")
            __check((describe(Wrapper.Empty())).toString(), "none")
        }
