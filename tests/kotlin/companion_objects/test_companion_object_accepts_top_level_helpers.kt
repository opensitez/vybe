// vybe-test: kotlin/companion_objects/test_companion_object_accepts_top_level_helpers
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

fun stampPrefix(value: String): String = "[" + value + "]"

        class Packet {
            companion object {
                fun label(value: String): String = stampPrefix(value)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Packet.label("x")).toString(), "[x]")
        }
