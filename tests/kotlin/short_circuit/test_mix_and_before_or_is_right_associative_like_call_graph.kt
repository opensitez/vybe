// vybe-test: kotlin/short_circuit/test_mix_and_before_or_is_right_associative_like_call_graph
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun l(v: String): Boolean {
                log += v
                return v == "go"
            }
            __check((l("go") && l("and") || l("tail")).toString(), "true")
            __check((log).toString(), "goand")
        }
