// vybe-test: kotlin/short_circuit/test_logical_and_skips_right_when_false
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun right(): Boolean { log += "right"
return true }
            __check((false && right()).toString(), "false")
            __check((log).toString(), "")
        }
