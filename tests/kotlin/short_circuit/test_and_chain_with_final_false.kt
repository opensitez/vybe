// vybe-test: kotlin/short_circuit/test_and_chain_with_final_false
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
return true }
            fun b(): Boolean { log += "b"
return true }
            fun c(): Boolean { log += "c"
return false }
            __check((a() && b() && c()).toString(), "false")
            __check((log).toString(), "abc")
        }
