// vybe-test: kotlin/function_overloads/test_overload_on_member_reference
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

class Ops {
            fun resolve(v: Int): String = "i"
            fun resolve(v: String): String = "s"
            fun use(): String {
                val fInt = this::resolve
                return fInt(3) + "," + fInt("x")
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Ops().use()).toString(), "i,s")
        }
