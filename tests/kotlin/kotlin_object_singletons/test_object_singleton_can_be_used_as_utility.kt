// vybe-test: kotlin/kotlin_object_singletons/test_object_singleton_can_be_used_as_utility
// origin: languages/kotlin/tests/kotlin/test_kotlin_object_singletons.rs

object Formatter {
            fun wrap(value: String): String = "<" + value + ">"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Formatter.wrap("a")).toString(), "<a>")
            __check((Formatter.wrap("b")).toString(), "<b>")
        }
