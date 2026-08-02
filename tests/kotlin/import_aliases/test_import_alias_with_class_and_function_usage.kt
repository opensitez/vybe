// vybe-test: kotlin/import_aliases/test_import_alias_with_class_and_function_usage
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.round as roundFunction

        fun normalize(v: Double): Int {
            return roundFunction(v).toInt()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((normalize(2.4)).toString(), "2")
            __check((roundFunction(2.9).toInt()).toString(), "3")
        }
