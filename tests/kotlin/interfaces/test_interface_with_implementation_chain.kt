// vybe-test: kotlin/interfaces/test_interface_with_implementation_chain
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Speaker {
            fun speak(): String
        }

        interface LoudSpeaker : Speaker {
            override fun speak(): String {
                return "loud"
            }
        }

        class Alarm : LoudSpeaker

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Alarm()
            __check((a.speak()).toString(), "loud")
        }
