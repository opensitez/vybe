// vybe-test: kotlin/kotlin_interface_defaults/test_interface_default_method_is_used_when_not_overridden
// origin: languages/kotlin/tests/kotlin/test_kotlin_interface_defaults.rs

interface Logger {
            fun prefix(): String = "log"
            fun format(message: String): String = prefix() + ":" + message
        }

        class DefaultLogger : Logger

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((DefaultLogger().format("ok")).toString(), "log:ok")
        }
