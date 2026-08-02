// vybe-test: kotlin/function_overloads/test_overload_in_generic_class_method
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

class Box {
            fun size(value: Int): String = "int"
            fun size(value: String): String = "str"
            fun size(value: List<Int>): String = "list"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.size(4)).toString(), "int")
            __check((b.size("a")).toString(), "str")
            __check((b.size(listOf(1))).toString(), "list")
        }
