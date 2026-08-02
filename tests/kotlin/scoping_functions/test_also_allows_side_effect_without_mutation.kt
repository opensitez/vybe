// vybe-test: kotlin/scoping_functions/test_also_allows_side_effect_without_mutation
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Logger {
            val events = mutableListOf<String>()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Logger().also {
                it.events.add("created")
                it.events.add("ready")
            }
            __check((value.events.joinToString(",")).toString(), "created,ready")
        }
