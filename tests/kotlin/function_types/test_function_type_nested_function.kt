// vybe-test: kotlin/function_types/test_function_type_nested_function
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun produce(): (String) -> Int {
                val base = 1
                return { it.length + base }
            }
            val f = produce()
            __check((f("ab")).toString(), "3")
        }
