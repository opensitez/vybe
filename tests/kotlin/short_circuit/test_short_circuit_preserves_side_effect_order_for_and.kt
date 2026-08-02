// vybe-test: kotlin/short_circuit/test_short_circuit_preserves_side_effect_order_for_and
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun a(): Boolean { log += "1"
return true }
            fun b(): Boolean { log += "2"
return false }
            __check((a() && b()).toString(), "false")
            __check((log).toString(), "12")
        }
