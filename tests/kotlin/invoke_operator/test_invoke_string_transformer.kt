// vybe-test: kotlin/invoke_operator/test_invoke_string_transformer
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Repeater {
            operator fun invoke(v: String, n: Int): String = v.repeat(n)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Repeater()("a", 3)).toString(), "aaa")
        }
