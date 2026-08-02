// vybe-test: kotlin/function_types/test_function_type_nullable
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun handle(fn: ((Int) -> Int)?): Int {
            return fn?.invoke(4) ?: 0
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((handle(null)).toString(), "0")
            __check((handle({ it + 1 })).toString(), "5")
        }
