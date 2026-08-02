// vybe-test: kotlin/short_circuit/test_when_with_guard_uses_short_circuit_or
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5
            var log = ""
            val out = when (value) {
                1, 2, 3 -> { log += "a"
1 }
                in 4..10 -> { log += "b"
2 }
                else -> { log += "c"
3 }
            }
            __check((out).toString(), "2")
            __check((log).toString(), "b")
        }
