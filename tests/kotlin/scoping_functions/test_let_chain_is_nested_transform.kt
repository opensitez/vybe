// vybe-test: kotlin/scoping_functions/test_let_chain_is_nested_transform
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
                .let { it + 1 }
                .let { it * 10 }
            __check((value).toString(), "40")
        }
