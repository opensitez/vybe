// vybe-test: kotlin/function_types/test_function_type_as_return_type
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun maker(): (Int) -> Int {
            return { v -> v * v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maker()(3)).toString(), "9")
        }
