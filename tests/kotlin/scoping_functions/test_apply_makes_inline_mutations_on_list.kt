// vybe-test: kotlin/scoping_functions/test_apply_makes_inline_mutations_on_list
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = mutableListOf(1).apply {
                add(2)
                add(3)
            }
            __check((list.joinToString("|")).toString(), "1|2|3")
            __check((list.size).toString(), "3")
        }
