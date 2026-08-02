// vybe-test: kotlin/this_super/test_this_in_require_context
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class Base {
        fun check() = "ok"
    }
    class Child : Base() {
        fun run() = this.check()
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Child().run()).toString(), "ok") }
