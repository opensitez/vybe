// vybe-test: kotlin/delegated_properties/test_observable_with_no_change
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

import kotlin.properties.Delegates

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val events = mutableListOf<String>()
            var value by Delegates.observable(10) { _, old, new ->
                events.add("${'$'}old/${'$'}new")
            }
            value = 12
            value = 12
            __check((events.size).toString(), "2")
            __check((events[0]).toString(), "10/12")
            __check((events[1]).toString(), "12/12")
        }
