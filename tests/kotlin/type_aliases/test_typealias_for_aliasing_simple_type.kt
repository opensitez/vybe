// vybe-test: kotlin/type_aliases/test_typealias_for_aliasing_simple_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Text = String

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Text = "ok"
            __check((value).toString(), "ok")
        }
