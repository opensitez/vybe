// vybe-test: kotlin/enums/test_enum_value_of_lookup
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class State { START, RUN, STOP }

        fun main() {
            println(State.valueOf("RUN"))
            try {
                State.valueOf("PAUSE")
                println("found")
            } catch (e: IllegalArgumentException) {
                println("missing")
            }
        }

