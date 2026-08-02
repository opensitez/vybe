// vybe-test: kotlin/function_types/test_function_type_stored_in_collection
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ops: List<(Int) -> Int> = listOf({ it + 1 }, { it * 2 }, { it * it })
            __check((ops[0](3)).toString(), "4")
            __check((ops[1](3)).toString(), "6")
            __check((ops[2](3)).toString(), "9")
        }
