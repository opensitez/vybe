// vybe-test: kotlin/short_circuit/test_safe_call_does_not_follow_false_branch
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun sideEffect(v: String): String { log += v
return v }
            val x: String? = null
            __check(((x == null) && (sideEffect("bad") == "bad")).toString(), "true")
            __check((log).toString(), "")
            val y: String? = "ok"
            __check(((y != null) || (sideEffect("bad") == "bad")).toString(), "true")
            __check((log).toString(), "")
        }
