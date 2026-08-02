// vybe-test: kotlin/type_aliases/test_typealias_for_receiver_function_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Formatter = String.() -> String

        val shout: Formatter = { this.uppercase() + "!" }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("k".shout()).toString(), "K!")
        }
