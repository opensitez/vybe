// vybe-test: kotlin/default_arguments/test_default_arguments_generic_default
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun <T> wrap(value: T, marker: String = "#"): String {
            return marker + value.toString()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((wrap(3)).toString(), "#3")
            __check((wrap("a", marker = "@")).toString(), "@a")
        }
