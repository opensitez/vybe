// vybe-test: kotlin/function_types/test_function_type_in_constructor
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

class Processor(val op: (Int) -> Int) {
            fun run(v: Int): Int = op(v)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Processor { it + 4 }.run(2)).toString(), "6")
        }
