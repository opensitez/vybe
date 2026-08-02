// vybe-test: kotlin/result_patterns/test_result_fold_failure
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { throw Exception("no") }
            val output = value.fold({ "s" + it.toString() }, { "f" })
            __check((output).toString(), "f")
        }
