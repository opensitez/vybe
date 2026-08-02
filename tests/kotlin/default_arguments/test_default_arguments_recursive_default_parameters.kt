// vybe-test: kotlin/default_arguments/test_default_arguments_recursive_default_parameters
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun depth(level: Int, suffix: String = ":") : String {
            return if (level <= 0) "0" else depth(level - 1, suffix) + suffix
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((depth(0)).toString(), "0")
            __check((depth(2)).toString(), "0::")
        }
