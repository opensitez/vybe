// vybe-test: kotlin/scoping_functions/test_scoping_also_for_logging_without_mutation
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 3
            val result = base.let { it * 2 }
                .also { }
            __check((result).toString(), "6")
        }
