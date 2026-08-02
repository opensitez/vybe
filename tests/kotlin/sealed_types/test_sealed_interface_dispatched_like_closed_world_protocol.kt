// vybe-test: kotlin/sealed_types/test_sealed_interface_dispatched_like_closed_world_protocol
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed interface Transport

        class Bus : Transport
        class Train : Transport

        fun is_mass_transport(value: Transport): Boolean {
            return when (value) {
                is Bus -> true
                is Train -> true
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((is_mass_transport(Bus())).toString(), "true")
            __check((is_mass_transport(Train())).toString(), "true")
        }
