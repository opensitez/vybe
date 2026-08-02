// vybe-test: kotlin/sealed_types/test_sealed_interface_like_hierarchy
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed interface Event

        class Start : Event
        class Stop : Event

        fun label(event: Event): String {
            return when (event) {
                is Start -> "start"
                is Stop -> "stop"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(Start())).toString(), "start")
            __check((label(Stop())).toString(), "stop")
        }
