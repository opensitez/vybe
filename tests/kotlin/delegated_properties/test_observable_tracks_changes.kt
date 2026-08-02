// vybe-test: kotlin/delegated_properties/test_observable_tracks_changes
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
            var value by Delegates.observable("init") { _, old, new ->
                events.add(old + ">" + new)
            }
            value = "a"
            value = "b"
            __check((value).toString(), "b")
            __check((events.joinToString(",")).toString(), "init>a,a>b")
        }
