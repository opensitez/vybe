// vybe-test: kotlin/named_arguments/test_named_arguments_called_from_defaulted_fn
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun greet(prefix: String, name: String, suffix: String = "!"): String {
            return prefix + name + suffix
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((greet(name = "k", prefix = "<", suffix = ">")).toString(), "<k>")
            __check((greet("[", "m")).toString(), "[m!")
        }
