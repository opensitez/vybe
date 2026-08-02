// vybe-test: kotlin/delegated_properties/test_observable_can_record_multiple_types
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

import kotlin.properties.Delegates

        class Tracker {
            var events = 0
            var label by Delegates.observable("x") { _, old, new ->
                if (old != new) events += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Tracker()
            t.label = "a"
            t.label = "a"
            t.label = "b"
            __check((t.label).toString(), "b")
            __check((t.events).toString(), "2")
        }
