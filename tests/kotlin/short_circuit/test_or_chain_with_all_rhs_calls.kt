// vybe-test: kotlin/short_circuit/test_or_chain_with_all_rhs_calls
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun a(): Boolean { log += "a"
return false }
            fun b(): Boolean { log += "b"
return false }
            fun c(): Boolean { log += "c"
return true }
            __check((a() || b() || c()).toString(), "true")
            __check((log).toString(), "abc")
        }
