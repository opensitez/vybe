// vybe-test: kotlin/function_types/test_function_type_through_map_entry
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map: Map<String, (Int) -> Int> = mapOf(
                "a" to { it + 1 },
                "b" to { it * 2 }
            )
            __check((map["a"]?.invoke(3)).toString(), "4")
            __check((map["b"]?.invoke(3)).toString(), "6")
        }
