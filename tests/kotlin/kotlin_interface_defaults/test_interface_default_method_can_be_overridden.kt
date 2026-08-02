// vybe-test: kotlin/kotlin_interface_defaults/test_interface_default_method_can_be_overridden
// origin: languages/kotlin/tests/kotlin/test_kotlin_interface_defaults.rs

interface Messenger {
            fun prefix(): String = "base"
            fun format(value: String): String = prefix() + ":" + value
        }

        class DefaultMessenger : Messenger

        class LoudMessenger : Messenger {
            override fun format(value: String): String = prefix() + ":" + value.toUpperCase()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((DefaultMessenger().format("ok")).toString(), "base:ok")
            __check((LoudMessenger().format("ok")).toString(), "base:OK")
        }
