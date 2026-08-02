// vybe-test: kotlin/data_class_copying/test_data_class_copy_overwrites_component_position
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Pair(val left: String, val right: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Pair("a", "b")
            val next = base.copy("x", right = "y")
            __check((next.left).toString(), "x")
            __check((next.right).toString(), "y")
        }
