// vybe-test: kotlin/invoke_operator/test_invoke_in_when_dispatch
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Router {
            operator fun invoke(flag: Boolean): String = if (flag) "ok" else "bad"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = Router()
            val a = true
            val b = false
            __check((r(a)).toString(), "ok")
            __check((r(b)).toString(), "bad")
        }
