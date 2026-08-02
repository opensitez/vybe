// vybe-test: kotlin/function_types/test_function_type_with_composition
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun comp(a: (Int) -> Int, b: (Int) -> Int): (Int) -> Int {
            return { x -> a(b(x)) }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = comp({ it + 1 }, { it * 2 })
            __check((f(3)).toString(), "7")
        }
