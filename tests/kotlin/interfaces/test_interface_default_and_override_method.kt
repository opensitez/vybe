// vybe-test: kotlin/interfaces/test_interface_default_and_override_method
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Messenger {
            fun send(message: String): String {
                return "default:" + message
            }
        }

        class Push : Messenger {
            override fun send(message: String): String {
                return "push:" + message
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Messenger = Push()
            __check((base.send("ok")).toString(), "push:ok")
        }
