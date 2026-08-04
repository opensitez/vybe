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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((describe(Wrapper.Item(Packet.Left(2)))).toString())
            __p((describe(Wrapper.Item(Packet.Right("ok")))).toString())
            __p((describe(Wrapper.Empty())).toString())
        
__check("L2\nRok\nnone")
}
