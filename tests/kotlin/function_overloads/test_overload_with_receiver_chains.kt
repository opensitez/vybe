// vybe-test: kotlin/function_overloads/test_overload_with_receiver_chains
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

class Maker {
            fun build(v: Int): Int = v
            fun build(v: Int, tag: String): String = tag + v
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = Maker().run {
                build(1).toString() + "," + build(1, "x")
            }
            __check((out).toString(), "1,x1")
        }
