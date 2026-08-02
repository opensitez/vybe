// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_nullable_coalesce
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val raw = listOf("10", "x", "20")
            val sum = raw.mapNotNull { it.toIntOrNull() }.sum()
            __check((sum).toString(), "30")
            __check((raw.mapNotNull { it.toIntOrNull(16) }.size).toString(), "0")
        }
