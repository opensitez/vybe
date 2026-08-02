// vybe-test: kotlin/data_class_destructuring/test_destructure_with_default_values_is_explicit_constructor
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Node(val value: Int = 4, val name: String = "x")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = Node(name = "n")
            val two = Node(3, "m")
            val (a, b) = one
            val (c, d) = two
            __check((a).toString(), "4")
            __check((b).toString(), "n")
            __check((c).toString(), "3")
            __check((d).toString(), "m")
        }
