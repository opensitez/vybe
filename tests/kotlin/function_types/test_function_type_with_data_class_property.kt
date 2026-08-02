// vybe-test: kotlin/function_types/test_function_type_with_data_class_property
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

data class Worker(val transform: (Int) -> Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w = Worker { it * it }
            __check((w.transform(5)).toString(), "25")
        }
