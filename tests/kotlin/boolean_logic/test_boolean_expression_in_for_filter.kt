// vybe-test: kotlin/boolean_logic/test_boolean_expression_in_for_filter
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "bb", "ccc")
            val ok = values.filter { it.length > 1 && it.length < 3 }
            __check((ok.joinToString(",")).toString(), "bb")
            val fail = values.filter { !ok.contains(it) }
            __check((fail.joinToString(",")).toString(), "a,ccc")
        }
