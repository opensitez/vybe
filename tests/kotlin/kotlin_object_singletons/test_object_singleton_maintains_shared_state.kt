// vybe-test: kotlin/kotlin_object_singletons/test_object_singleton_maintains_shared_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_object_singletons.rs

object Config {
            var enabled = false
            fun enable() { enabled = true }
            fun isEnabled(): Boolean = enabled
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Config.enable()
            __check((Config.isEnabled()).toString(), "true")
            Config.enabled = false
            __check((Config.isEnabled()).toString(), "false")
        }
