// vybe-test: kotlin/short_circuit/test_mix_or_before_and_respects_grouping
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun l(v: String): Boolean { log += v
return v == "go" }
            __check(((l("go") || l("skip")) && l("and")).toString(), "true")
            __check((log).toString(), "go")
        }
