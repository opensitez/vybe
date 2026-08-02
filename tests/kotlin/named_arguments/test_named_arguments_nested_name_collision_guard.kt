// vybe-test: kotlin/named_arguments/test_named_arguments_nested_name_collision_guard
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun combine(a: String, b: String): String = a + b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = "x"
            __check((combine(a = "u", b = a)).toString(), "ux")
        }
