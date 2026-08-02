// vybe-test: kotlin/constructor_chaining/test_constructor_with_interface_impl
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

interface I { fun tag(): String }
        class C(val v: Int) : I {
            override fun tag() = "c:$v"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((C(7).tag()).toString(), "c:7")
        }
