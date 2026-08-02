// vybe-test: kotlin/literals/test_empty_and_null_coalescing_literal_source
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nullable: String? = null
            __check((nullable?.let { it.length } ?: -1).toString(), "-1")
            val present: String? = "k"
            __check((present?.length ?: 0).toString(), "1")
        }
