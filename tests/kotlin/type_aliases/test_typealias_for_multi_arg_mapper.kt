// vybe-test: kotlin/type_aliases/test_typealias_for_multi_arg_mapper
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Joiner = (String, String, String) -> String

        fun join(parts: Joiner): String {
            return parts("a", "b", "c")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val joiner: Joiner = { left, middle, right -> left + "-" + middle + "-" + right }
            __check((join(joiner)).toString(), "a-b-c")
        }
