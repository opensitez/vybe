// vybe-test: kotlin/function_types/test_function_type_of_unit
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun apply(side: (Int) -> Unit): String {
            side(3)
            return "ok"
        }
        fun main() {
            println(apply({ println("x" + it); }))
        }

