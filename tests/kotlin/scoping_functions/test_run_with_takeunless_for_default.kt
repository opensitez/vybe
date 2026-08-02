// vybe-test: kotlin/scoping_functions/test_run_with_takeunless_for_default
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val candidate = "x".takeUnless { it.isEmpty() } ?: "missing"
            val result = candidate.run {
                "found:" + this
            }
            __check((result).toString(), "found:x")
        }
