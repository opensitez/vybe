// vybe-test: kotlin/boolean_logic/test_boolean_guarded_mapping
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = listOf(1, 2, 3, 4, 5)
            val filtered = input.filter { it > 2 && it < 5 }
            val mapped = filtered.map { it % 2 == 0 }
            __check((filtered.joinToString(",")).toString(), "3,4")
            __check((mapped.joinToString(",")).toString(), "false,true")
            __check((filtered.size).toString(), "2")
        }
