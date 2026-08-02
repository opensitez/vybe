// vybe-test: kotlin/object_declarations/test_object_properties_can_reference_companion
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Config {
            val enabled = true
        }

        class Processor {
            fun active(): String = if (Config.enabled) "yes" else "no"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Processor().active()).toString(), "yes")
        }
