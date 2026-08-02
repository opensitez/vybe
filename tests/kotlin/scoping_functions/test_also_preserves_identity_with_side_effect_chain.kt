// vybe-test: kotlin/scoping_functions/test_also_preserves_identity_with_side_effect_chain
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = mutableListOf(1)
            val second = first
                .also { it.add(2) }
                .also { it.add(3) }
            __check((first === second).toString(), "true")
            __check((second.joinToString("|")).toString(), "1|2|3")
        }
