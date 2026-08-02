// vybe-test: kotlin/constructor_chaining/test_constructor_multiple_init
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Multi {
            val a: Int
            init {
                __check(("init").toString(), "init")
                a = 1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Multi()
            __check((m.a).toString(), "1")
        }
