// vybe-test: kotlin/kotlin_lazy_delegates/test_delegates_observable_records_changes
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

import kotlin.properties.Delegates

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var events = ""
            var value by Delegates.observable(1) { _, old, new ->
                events += old.toString() + ":" + new.toString() + ";"
            }
            value = 2
            value = 5
            __check((events).toString(), "1:2;2:5;")
            __check((value).toString(), "5")
        }
