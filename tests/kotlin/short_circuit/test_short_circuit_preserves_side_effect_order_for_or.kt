// vybe-test: kotlin/short_circuit/test_short_circuit_preserves_side_effect_order_for_or
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
return false }
            fun b(): Boolean { log += "2"
return true }
            __check((a() || b()).toString(), "true")
            __check((log).toString(), "12")
        }
